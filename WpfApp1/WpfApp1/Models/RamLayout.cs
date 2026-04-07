namespace WpfApp1.Models
{
    public class RamLayout
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public string Address { get; set; } = string.Empty;
        public int Offset { get; set; }
        public string Symbol { get; set; } = string.Empty;
        public string FormatId { get; set; } = string.Empty;
    }
}
