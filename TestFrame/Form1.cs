using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TestFrame
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();


            Task.Run(() =>
            {
                while (true)
                {
                    Task.Delay(10);

                    using var currentFrame = new ScreenCapturerWin().GetNextFrame();
                    if (currentFrame == null)
                    {
                        return;
                    }

                    pictureBox1.InvokeIfRequired(() =>
                    {
                        pictureBox1.Image = Image.FromHbitmap(currentFrame.GetHbitmap());
                        pictureBox1.Show();
                        pictureBox1.Refresh();
                    });

                }
            });
        }
    }

    public static class Extension
    {
        public static void InvokeIfRequired(this Control control, MethodInvoker action)
        {
            //在非當前執行緒內 使用委派
            if (control.InvokeRequired)
            {
                control.Invoke(action);
            }
            else
            {
                action();
            }
        }
    }
}