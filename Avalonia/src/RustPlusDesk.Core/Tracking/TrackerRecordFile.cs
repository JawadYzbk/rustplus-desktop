using System.Buffers.Binary;
using System.Globalization;
using System.Text;

namespace RustPlusDesk.Core.Tracking;

/// <summary>
/// The on-disk layout of one player's observations in one wipe, and what is needed to keep appending to it.
///
/// It replaced JSON lines. A JSON record spelled out every field name, the Steam id, a 32-character session id and a
/// 28-character timestamp on every poll, which came to about 390 bytes; this comes to about 30. Strings that repeat
/// - the player name, the session id, the monument name, the grid - are written once as a definition record and
/// referred to by number afterwards.
///
/// Layout: "PWT1", then the base timestamp in ticks (int64, little endian), then records:
///   1  string       varint id, varint byte length, UTF-8 bytes
///   2  observation  zigzag ticks since the base, flags, location type, varint references for the
///                   session, the name, the location name and the grid (0 = none), float32 x/y when
///                   present, zigzag spawn and death time when present
///
/// Every record is self-delimiting and the timestamps are absolute against the header, so reading can start at any
/// record boundary once the string table is known. A record cut short by a crash ends the readable part of the file;
/// the next append truncates it away.
/// </summary>
public sealed class TrackerRecordFile
{
    public const int HeaderLength = 12;
    private static readonly byte[] Magic = "PWT1"u8.ToArray();
    private const byte StringTag = 1;
    private const byte ObservationTag = 2;

    [Flags]
    private enum RecordFlags : byte
    {
        Connected = 1,
        SnapshotValid = 2,
        Online = 4,
        Dead = 8,
        Afk = 16,
        HasPosition = 32,
        HasSpawnTime = 64,
        HasDeathTime = 128,
    }

    private readonly List<string> strings = new();
    private readonly Dictionary<string, int> ids = new(StringComparer.Ordinal);

    private TrackerRecordFile(long baseTicks, long validLength)
    {
        BaseTicks = baseTicks;
        ValidLength = validLength;
    }

    /// <summary>The timestamp every record's offset is counted from.</summary>
    public long BaseTicks { get; }

    /// <summary>Where the last complete record ends; anything after it is a torn write.</summary>
    public long ValidLength { get; private set; }

    public static TrackerRecordFile CreateNew(DateTime firstTimestampUtc) =>
        new(firstTimestampUtc.ToUniversalTime().Ticks, HeaderLength);

    /// <summary>The header of a file: the magic and the base timestamp.</summary>
    public static byte[] Header(long baseTicks)
    {
        var header = new byte[HeaderLength];
        Magic.CopyTo(header, 0);
        BinaryPrimitives.WriteInt64LittleEndian(header.AsSpan(4), baseTicks);
        return header;
    }

    /// <summary>Whether the bytes start with this format's header.</summary>
    public static bool HasHeader(ReadOnlySpan<byte> bytes) =>
        bytes.Length >= HeaderLength && bytes[..Magic.Length].SequenceEqual(Magic);

    /// <summary>Reads a whole file, building the string table; null when it is not this format.</summary>
    public static TrackerRecordFile? Parse(byte[] bytes, ulong steamId, List<PlayerObservation>? sink)
    {
        if (!HasHeader(bytes)) return null;
        var file = new TrackerRecordFile(BinaryPrimitives.ReadInt64LittleEndian(bytes.AsSpan(4)), HeaderLength);
        file.ValidLength = HeaderLength + file.Read(bytes.AsSpan(HeaderLength), steamId, sink);
        return file;
    }

    /// <summary>
    /// Decodes the records from a slice that starts on a record boundary and returns how many bytes were complete.
    /// The strings defined in the slice join the table, so the same instance can keep reading what is appended later.
    /// </summary>
    public int Read(ReadOnlySpan<byte> bytes, ulong steamId, List<PlayerObservation>? sink)
    {
        var position = 0;
        var complete = 0;
        while (position < bytes.Length)
        {
            if (!TryReadRecord(bytes, ref position, steamId, sink)) break;
            complete = position;
        }
        return complete;
    }

    /// <summary>Encodes one observation, with the string definitions it needs ahead of it.</summary>
    public byte[] Encode(PlayerObservation observation)
    {
        using var buffer = new MemoryStream(48);
        var session = Ref(observation.SessionId, buffer);
        var name = Ref(observation.Name, buffer);
        var locationName = Ref(observation.LocationName, buffer);
        var grid = Ref(observation.Grid, buffer);

        var flags = (observation.IsConnected ? RecordFlags.Connected : 0)
            | (observation.SnapshotValid ? RecordFlags.SnapshotValid : 0)
            | (observation.Online ? RecordFlags.Online : 0)
            | (observation.Dead ? RecordFlags.Dead : 0)
            | (observation.Afk ? RecordFlags.Afk : 0)
            | (observation.X.HasValue && observation.Y.HasValue ? RecordFlags.HasPosition : 0)
            | (observation.SpawnTime.HasValue ? RecordFlags.HasSpawnTime : 0)
            | (observation.DeathTime.HasValue ? RecordFlags.HasDeathTime : 0);

        buffer.WriteByte(ObservationTag);
        WriteVarint(buffer, ZigZag(observation.TimestampUtc.ToUniversalTime().Ticks - BaseTicks));
        buffer.WriteByte((byte)flags);
        buffer.WriteByte((byte)observation.LocationType);
        WriteVarint(buffer, (ulong)session);
        WriteVarint(buffer, (ulong)name);
        WriteVarint(buffer, (ulong)locationName);
        WriteVarint(buffer, (ulong)grid);
        if ((flags & RecordFlags.HasPosition) != 0)
        {
            // Rust+ sends the positions as float32, so this is lossless for everything recorded live.
            Span<byte> position = stackalloc byte[8];
            BinaryPrimitives.WriteSingleLittleEndian(position, (float)observation.X!.Value);
            BinaryPrimitives.WriteSingleLittleEndian(position[4..], (float)observation.Y!.Value);
            buffer.Write(position);
        }
        if (observation.SpawnTime is { } spawn) WriteVarint(buffer, ZigZag(spawn));
        if (observation.DeathTime is { } death) WriteVarint(buffer, ZigZag(death));
        return buffer.ToArray();
    }

    /// <summary>Records that the bytes produced by <see cref="Encode"/> reached the file.</summary>
    public void Appended(int length) => ValidLength += length;

    private int Ref(string? value, MemoryStream buffer)
    {
        if (value is null) return 0;
        if (ids.TryGetValue(value, out var id)) return id;

        id = strings.Count + 1;
        strings.Add(value);
        ids[value] = id;
        var bytes = Encoding.UTF8.GetBytes(value);
        buffer.WriteByte(StringTag);
        WriteVarint(buffer, (ulong)id);
        WriteVarint(buffer, (ulong)bytes.Length);
        buffer.Write(bytes);
        return id;
    }

    private bool TryReadRecord(ReadOnlySpan<byte> bytes, ref int position, ulong steamId, List<PlayerObservation>? sink)
    {
        var cursor = position;
        if (!TryReadByte(bytes, ref cursor, out var tag)) return false;

        if (tag == StringTag)
        {
            if (!TryReadVarint(bytes, ref cursor, out var id) || !TryReadVarint(bytes, ref cursor, out var length) ||
                length > int.MaxValue || cursor + (long)length > bytes.Length) return false;
            var value = Encoding.UTF8.GetString(bytes.Slice(cursor, (int)length));
            cursor += (int)length;
            // Reading again from an offset meets the definitions that are already in the table; only the next unseen id
            // extends it.
            if (id == (ulong)strings.Count + 1)
            {
                strings.Add(value);
                ids.TryAdd(value, (int)id);
            }
            else if (id > (ulong)strings.Count)
            {
                return false;
            }
            position = cursor;
            return true;
        }

        if (tag != ObservationTag) return false;

        if (!TryReadVarint(bytes, ref cursor, out var ticks) ||
            !TryReadByte(bytes, ref cursor, out var rawFlags) ||
            !TryReadByte(bytes, ref cursor, out var location) ||
            !TryReadVarint(bytes, ref cursor, out var session) ||
            !TryReadVarint(bytes, ref cursor, out var name) ||
            !TryReadVarint(bytes, ref cursor, out var locationName) ||
            !TryReadVarint(bytes, ref cursor, out var grid)) return false;

        var flags = (RecordFlags)rawFlags;
        double? x = null, y = null;
        if ((flags & RecordFlags.HasPosition) != 0)
        {
            if (cursor + 8 > bytes.Length) return false;
            x = BinaryPrimitives.ReadSingleLittleEndian(bytes[cursor..]);
            y = BinaryPrimitives.ReadSingleLittleEndian(bytes[(cursor + 4)..]);
            cursor += 8;
        }

        long? spawnTime = null, deathTime = null;
        if ((flags & RecordFlags.HasSpawnTime) != 0)
        {
            if (!TryReadVarint(bytes, ref cursor, out var spawn)) return false;
            spawnTime = UnZigZag(spawn);
        }
        if ((flags & RecordFlags.HasDeathTime) != 0)
        {
            if (!TryReadVarint(bytes, ref cursor, out var death)) return false;
            deathTime = UnZigZag(death);
        }

        var timestampTicks = BaseTicks + UnZigZag(ticks);
        if (timestampTicks < DateTime.MinValue.Ticks || timestampTicks > DateTime.MaxValue.Ticks) return false;

        sink?.Add(new PlayerObservation(
            steamId,
            Lookup(name) ?? steamId.ToString(CultureInfo.InvariantCulture),
            new DateTime(timestampTicks, DateTimeKind.Utc),
            Lookup(session) ?? string.Empty,
            (flags & RecordFlags.Connected) != 0,
            (flags & RecordFlags.SnapshotValid) != 0,
            (flags & RecordFlags.Online) != 0,
            (flags & RecordFlags.Dead) != 0,
            (flags & RecordFlags.Afk) != 0,
            x,
            y,
            Enum.IsDefined(typeof(TrackerLocationType), (int)location) ? (TrackerLocationType)location : TrackerLocationType.Unknown,
            Lookup(locationName),
            Lookup(grid),
            spawnTime,
            deathTime));

        position = cursor;
        return true;
    }

    private string? Lookup(ulong id) => id == 0 || id > (ulong)strings.Count ? null : strings[(int)id - 1];

    private static void WriteVarint(Stream stream, ulong value)
    {
        Span<byte> buffer = stackalloc byte[10];
        var length = 0;
        while (value >= 0x80)
        {
            buffer[length++] = (byte)(value | 0x80);
            value >>= 7;
        }
        buffer[length++] = (byte)value;
        stream.Write(buffer[..length]);
    }

    private static bool TryReadVarint(ReadOnlySpan<byte> bytes, ref int cursor, out ulong value)
    {
        value = 0;
        var shift = 0;
        while (cursor < bytes.Length)
        {
            var current = bytes[cursor++];
            value |= (ulong)(current & 0x7F) << shift;
            if ((current & 0x80) == 0) return true;
            shift += 7;
            if (shift > 63) return false;
        }
        return false;
    }

    private static bool TryReadByte(ReadOnlySpan<byte> bytes, ref int cursor, out byte value)
    {
        value = 0;
        if (cursor >= bytes.Length) return false;
        value = bytes[cursor++];
        return true;
    }

    private static ulong ZigZag(long value) => (ulong)((value << 1) ^ (value >> 63));

    private static long UnZigZag(ulong value) => (long)(value >> 1) ^ -(long)(value & 1);
}
