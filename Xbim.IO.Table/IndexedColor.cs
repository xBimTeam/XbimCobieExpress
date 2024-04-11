namespace Xbim.IO.Table
{
    public class IndexedColor
    {
        public static readonly IndexedColor Black;
        public static readonly IndexedColor White;
        public static readonly IndexedColor Red;
        public static readonly IndexedColor BrightGreen;
        public static readonly IndexedColor Blue;
        public static readonly IndexedColor Yellow;
        public static readonly IndexedColor Pink;
        public static readonly IndexedColor Turquoise;
        public static readonly IndexedColor DarkRed;
        public static readonly IndexedColor Green;
        public static readonly IndexedColor DarkBlue;
        public static readonly IndexedColor DarkYellow;
        public static readonly IndexedColor Violet;
        public static readonly IndexedColor Teal;
        public static readonly IndexedColor Grey25Percent;
        public static readonly IndexedColor Grey50Percent;
        public static readonly IndexedColor CornflowerBlue;
        public static readonly IndexedColor Maroon;
        public static readonly IndexedColor LemonChiffon;
        public static readonly IndexedColor Orchid;
        public static readonly IndexedColor Coral;
        public static readonly IndexedColor RoyalBlue;
        public static readonly IndexedColor LightCornflowerBlue;
        public static readonly IndexedColor SkyBlue;
        public static readonly IndexedColor LightTurquoise;
        public static readonly IndexedColor LightGreen;
        public static readonly IndexedColor LightYellow;
        public static readonly IndexedColor PaleBlue;
        public static readonly IndexedColor Rose;
        public static readonly IndexedColor Lavender;
        public static readonly IndexedColor Tan;
        public static readonly IndexedColor LightBlue;
        public static readonly IndexedColor Aqua;
        public static readonly IndexedColor Lime;
        public static readonly IndexedColor Gold;
        public static readonly IndexedColor LightOrange;
        public static readonly IndexedColor Orange;
        public static readonly IndexedColor BlueGrey;
        public static readonly IndexedColor Grey40Percent;
        public static readonly IndexedColor DarkTeal;
        public static readonly IndexedColor SeaGreen;
        public static readonly IndexedColor DarkGreen;
        public static readonly IndexedColor OliveGreen;
        public static readonly IndexedColor Brown;
        public static readonly IndexedColor Plum;
        public static readonly IndexedColor Indigo;
        public static readonly IndexedColor Grey80Percent;
        public static readonly IndexedColor Automatic;

        private int index;
        private byte[] _rgb;


        IndexedColor(int idx, byte[] rgb)
        {
            index = idx;
            _rgb = rgb;
        }

        static IndexedColor()
        {
            Black = new IndexedColor(8, new byte[] { 0, 0, 0 });
            White = new IndexedColor(9, new byte[] { 255, 255, 255 });
            Red = new IndexedColor(10, new byte[] { 255, 0, 0 });
            BrightGreen = new IndexedColor(11, new byte[] { 0, 255, 0 });
            Blue = new IndexedColor(12, new byte[] { 0, 0, 255 });
            Yellow = new IndexedColor(13, new byte[] { 255, 255, 0 });
            Pink = new IndexedColor(14, new byte[] { 255, 0, 255 });
            Turquoise = new IndexedColor(15, new byte[] { 0, 255, 255 });
            DarkRed = new IndexedColor(16, new byte[] { 128, 0, 0 });
            Green = new IndexedColor(17, new byte[] { 0, 128, 0 });
            DarkBlue = new IndexedColor(18, new byte[] { 0, 0, 128 });
            DarkYellow = new IndexedColor(19, new byte[] { 128, 128, 0 });
            Violet = new IndexedColor(20, new byte[] { 128, 0, 128 });
            Teal = new IndexedColor(21, new byte[] { 0, 128, 128 });
            Grey25Percent = new IndexedColor(22, new byte[] { 192, 192, 192 });
            Grey50Percent = new IndexedColor(23, new byte[] { 128, 128, 128 });
            CornflowerBlue = new IndexedColor(24, new byte[] { 153, 153, 255 });
            Maroon = new IndexedColor(25, new byte[] { 127, 0, 0 });
            LemonChiffon = new IndexedColor(26, new byte[] { 255, 255, 204 });
            Orchid = new IndexedColor(28, new byte[] { 102, 0, 102 });
            Coral = new IndexedColor(29, new byte[] { 255, 128, 128 });
            RoyalBlue = new IndexedColor(30, new byte[] { 0, 102, 204 });
            LightCornflowerBlue = new IndexedColor(31, new byte[] { 204, 204, 255 });
            SkyBlue = new IndexedColor(40, new byte[] { 0, 204, 255 });
            LightTurquoise = new IndexedColor(41, new byte[] { 204, 255, 255 });
            LightGreen = new IndexedColor(42, new byte[] { 204, 255, 204 });
            LightYellow = new IndexedColor(43, new byte[] { 255, 255, 153 });
            PaleBlue = new IndexedColor(44, new byte[] { 153, 204, 255 });
            Rose = new IndexedColor(45, new byte[] { 255, 153, 204 });
            Lavender = new IndexedColor(46, new byte[] { 204, 153, 255 });
            Tan = new IndexedColor(47, new byte[] { 255, 204, 153 });
            LightBlue = new IndexedColor(48, new byte[] { 51, 102, 255 });
            Aqua = new IndexedColor(49, new byte[] { 51, 204, 204 });
            Lime = new IndexedColor(50, new byte[] { 153, 204, 0 });
            Gold = new IndexedColor(51, new byte[] { 255, 204, 0 });
            LightOrange = new IndexedColor(52, new byte[] { 255, 153, 0 });
            Orange = new IndexedColor(53, new byte[] { 255, 102, 0 });
            BlueGrey = new IndexedColor(54, new byte[] { 102, 102, 153 });
            Grey40Percent = new IndexedColor(55, new byte[] { 150, 150, 150 });
            DarkTeal = new IndexedColor(56, new byte[] { 0, 51, 102 });
            SeaGreen = new IndexedColor(57, new byte[] { 51, 153, 102 });
            DarkGreen = new IndexedColor(58, new byte[] { 0, 51, 0 });
            OliveGreen = new IndexedColor(59, new byte[] { 51, 51, 0 });
            Brown = new IndexedColor(60, new byte[] { 153, 51, 0 });
            Plum = new IndexedColor(61, new byte[] { 153, 51, 102 });
            Indigo = new IndexedColor(62, new byte[] { 51, 51, 153 });
            Grey80Percent = new IndexedColor(63, new byte[] { 51, 51, 51 });
            Automatic = new IndexedColor(64, new byte[] { 0, 0, 0 }); 

        }
        public byte[] RGB
        {
            get { return _rgb; }
        }
       
        public short Index
        {
            get
            {
                return (short)index;
            }
        }
    }
}
