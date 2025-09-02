using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OOX_SVOD
{
    public static class Extensions
    {
        private delegate void TextDelegate(Control textBox, string text);


        public static void SetTextToControl(Control textBox, string text)
        {
            if (textBox.InvokeRequired)
            {
                var d = new TextDelegate(SetTextToControl);
                textBox.Invoke(d, new object[] { textBox, text });
            }
            else
                textBox.Text = text;
        }

        public static void AddTextToControl(Control textBox, string text)
        {
            if (textBox.InvokeRequired)
            {
                var d = new TextDelegate(SetTextToControl);
                textBox.Invoke(d, new object[] { textBox, text });
            }
            else
                textBox.Text = textBox.Text + text + "\r\n";
        }
    }
}
