namespace RustPlusDesk.Core.Cameras;

/// <summary>
/// Reference CameraWindow: the stick and the keys the camera takes, and where a player the camera sees is drawn.
/// </summary>
public static class CameraInput
{
    /// <summary>The stick's dead zone: below this the camera is not told to move.</summary>
    public const double DeadZone = 0.35;

    /// <summary>Reference JoyResponse: a fine control near the centre, fast toward the edge.</summary>
    public static float Response(double value, double speed)
    {
        double a = Math.Abs(value);
        if (a < 0.08) return 0f;
        return (float)(Math.Sign(value) * a * a * speed);
    }

    /// <summary>
    /// Reference JoyUpdate: the stick's position, clamped to the circle, with the right positive and the up positive.
    /// </summary>
    public static (double X, double Y) Stick(double dx, double dy, double maxRadius)
    {
        double len = Math.Sqrt(dx * dx + dy * dy);
        if (len > maxRadius && len > 0)
        {
            dx = dx / len * maxRadius;
            dy = dy / len * maxRadius;
        }
        return (dx / maxRadius, -dy / maxRadius);
    }

    /// <summary>Reference JoystickTick: the keys the stick sends.</summary>
    public static CameraButtons MovementButtons(double x, double y)
    {
        var buttons = CameraButtons.None;
        if (y > DeadZone) buttons |= CameraButtons.Forward;
        else if (y < -DeadZone) buttons |= CameraButtons.Backward;
        if (x > DeadZone) buttons |= CameraButtons.Right;
        else if (x < -DeadZone) buttons |= CameraButtons.Left;
        return buttons;
    }
}

/// <summary>Reference CameraWindow.DrawOverlay: where a player the camera sees is drawn in the picture.</summary>
public static class CameraOverlay
{
    /// <summary>
    /// The place of an entity in the picture, or null when it is behind the camera or outside the frame. The
    /// reference lifts the name a little above the player, so the numbers here are the player's place and the caller
    /// lifts the name.
    /// </summary>
    public static (double X, double Y)? Place(CameraEntity entity, double verticalFovDegrees, int width, int height)
    {
        if (width <= 0 || height <= 0) return null;
        double z = entity.Z;
        if (z <= 0.01) return null;

        double vf = verticalFovDegrees * Math.PI / 180.0;
        double aspect = width / (double)height;
        double hf = 2.0 * Math.Atan(Math.Tan(vf / 2.0) * aspect);
        double xndc = (entity.X / z) / Math.Tan(hf / 2.0);
        double yndc = (entity.Y / z) / Math.Tan(vf / 2.0);
        double u = (xndc * 0.5 + 0.5) * width;
        double v = (-yndc * 0.5 + 0.5) * height;
        if (u < -10 || u > width + 10 || v < -10 || v > height + 10) return null;
        return (u, v);
    }
}
