using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Interface
{
    public partial class ImagePreview : Form
    {
        public ImagePreview()
        {
            InitializeComponent();
            pictureBox1.Image = Clipboard.GetImage();
            this.Size = pictureBox1.Size;
        }
    }
}
