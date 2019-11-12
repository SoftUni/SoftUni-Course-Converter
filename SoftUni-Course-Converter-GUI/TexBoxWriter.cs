using System.IO;
using System.Text;
using System.Windows.Forms;

namespace SoftUni_Course_Converter
{
    public class TexBoxWriter : TextWriter
    {
        private TextBox textBox;

        public TexBoxWriter(TextBox textBox)
        {
            this.textBox = textBox;
        }

        public override Encoding Encoding => Encoding.UTF8;

        public override void WriteLine(string text)
        {
            this.textBox.AppendText(text + "\r\n");
        }

        public override void WriteLine(string format, params object[] args)
        {
            this.WriteLine(string.Format(format, args));
        }

        public override void WriteLine()
        {
            this.WriteLine("");
        }
    }
}