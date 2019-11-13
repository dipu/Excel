using System.ComponentModel;
using System.Drawing;

namespace Dipu.Excel
{
    public class Style
    {
        public Style(Color backgroundColor, Color textColor)
        {
            BackgroundColor = backgroundColor;
            TextColor = textColor;
        }

        public Color BackgroundColor { get; }

        public Color TextColor { get; }
    }
}
