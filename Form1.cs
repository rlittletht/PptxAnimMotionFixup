using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Permissions;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace FixPPTLayout
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void DoFixup(object sender, EventArgs e)
        {
            PresentationX.Slide sld = (PresentationX.Slide)m_cbSlides.SelectedItem;

            if (System.IO.File.Exists(m_ebOutput.Text))
                System.IO.File.Delete(m_ebOutput.Text);

            System.IO.File.Copy(m_eb169.Text, m_ebOutput.Text);
            PresentationX pptx = new PresentationX(m_ebOutput.Text, FileMode.Open, FileAccess.ReadWrite, FileShare.None);

            pptx.FixSlideAnimMotion(sld, sld.Uri);
            pptx.Close();
            MessageBox.Show("Done!", "FixPPTAnimation");
        }

        private PresentationX m_pptx;

        private void DoOpenFiles(object sender, EventArgs e)
        {
            m_pptx = new PresentationX(m_eb43.Text, FileMode.Open, FileAccess.Read, FileShare.Read);

            List<PresentationX.Slide> plsld = m_pptx.GetSlides();

            foreach (PresentationX.Slide sld in plsld)
                {
                m_cbSlides.Items.Add(sld);
                }
        }

        private void DoFixAll(object sender, EventArgs e)
        {
            if (System.IO.File.Exists(m_ebOutput.Text))
                System.IO.File.Delete(m_ebOutput.Text);

            System.IO.File.Copy(m_eb169.Text, m_ebOutput.Text);
            PresentationX pptx = new PresentationX(m_ebOutput.Text, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            foreach (PresentationX.Slide sld in m_cbSlides.Items)
                {
                pptx.FixSlideAnimMotion(sld, sld.Uri);
                }
            pptx.Close();
            MessageBox.Show("Done!", "FixPPTAnimation");

        }
    }
}
