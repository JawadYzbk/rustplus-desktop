namespace RustPlusDesk.Core.Cloud;

/// <summary>Verified plan facts relevant to cloud device synchronization; never inferred from a tier name.</summary>
public sealed record CloudPlan(string Code, bool IsPremium, bool CanSync, int? MaxDevices);
