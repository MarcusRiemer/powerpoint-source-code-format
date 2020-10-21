using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools.Ribbon;

namespace pp_source_format
{
    public partial class Ribbon
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            CheckPygmentizeStatus();
        }

        private void CheckPygmentizeStatus()
        {
            try
            {
                var PygmentizePath = Formatter.FindPygmentizePath();
                lblPygmentsAvailable.SuperTip = PygmentizePath;
                
                foreach (var c in InActivePygmentizeControls)
                {
                    c.Visible = false;
                }

                foreach (var c in ActivePygmentizeControls)
                {
                    c.Enabled = true;
                }

                lblPygmentsAvailable.Visible = true;
                lblPygmentsNotAvailable.Visible = false;
            }
            catch (Exception ex)
            {
                foreach (var c in InActivePygmentizeControls)
                {
                    c.Visible = true;
                }

                foreach (var c in ActivePygmentizeControls)
                {
                    c.Enabled = false;
                }


                lblPygmentsAvailable.Visible = false;
                lblPygmentsNotAvailable.Visible = true;
            }
        }

        private void OnRenameSelected(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var selectedShapes = ActiveWindow.Selection.ShapeRange;
                foreach (Shape shape in selectedShapes)
                {
                    Formatter.FormatShape(shape, cmbLanguage.Text, cmbStyle.Text);
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message, ex.GetType().Name);
            }
        }

        private Application Application
        {
            get => Globals.SourceCodeFormatAddin.Application;
        }

        private Presentation ActivePresentation
        {
            get => Application.ActivePresentation;
        }

        private DocumentWindow ActiveWindow
        {
            get => Application.ActiveWindow;
        }

        private IEnumerable<RibbonControl> ActivePygmentizeControls
        {
            get => new RibbonControl[] { lblPygmentsAvailable, btnFormatAll, btnFormatCurrent, cmbLanguage, };
        }

        private IEnumerable<RibbonControl> InActivePygmentizeControls
        {
            get => new RibbonControl[] { lblPygmentsNotAvailable, btnHelpPygmentize };
        }
    }
}
