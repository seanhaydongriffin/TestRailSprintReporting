
namespace SharedProject
{
    public static class Color
    {
        public static string NameToHex(string name)
        {
            int ColorValue = System.Drawing.Color.FromName(name).ToArgb() & 0xFFFFFF;
            return string.Format("#{0:X6}", ColorValue);
        }
    }
}
