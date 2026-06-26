namespace RustPlusDesk.Services.InventoryScan
{
    public sealed class RecognizedInventoryStack
    {
        public string ShortName { get; set; } = "";
        public string DisplayName { get; set; } = "";
        public int Quantity { get; set; }
        public double IconConfidence { get; set; }
        public int SlotIndex { get; set; }
    }
}