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
            ReflectPygmentizeStatus();
        }

        /// <summary>
        /// As it is not exactly straightforwad to get hold of pygmentize the UI
        /// tries to give at least a little feedback.
        /// </summary>
        private void ReflectPygmentizeStatus()
        {
            try
            {
                // This will throw if pygmentize is not available
                var PygmentizePath = Formatter.FindPygmentizePath();

                // So here we are on the success path
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
                // And here we disable / hide the operations that are not meaningful
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

        /// <summary>
        /// The user has decided to format some shapes
        /// </summary>
        private void OnFormatSelected(object sender, RibbonControlEventArgs e)
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

        /// <summary>
        /// Shortcut to global Powerpoint object
        /// </summary>
        private Application Application
        {
            get => Globals.SourceCodeFormatAddin.Application;
        }

        /// <summary>
        /// Shortcut to the active presentation
        /// </summary>
        private Presentation ActivePresentation
        {
            get => Application.ActivePresentation;
        }

        /// <summary>
        /// Shortcut to the active Window
        /// </summary>
        private DocumentWindow ActiveWindow
        {
            get => Application.ActiveWindow;
        }

        /// <summary>
        /// All controls that are meaningful if pygmentize was found
        /// </summary>
        private IEnumerable<RibbonControl> ActivePygmentizeControls
        {
            get => new RibbonControl[] { lblPygmentsAvailable, btnFormatAll, btnFormatCurrent, cmbLanguage, };
        }

        /// <summary>
        /// All controls that are meaningful if pygmentize is missing
        /// </summary>
        private IEnumerable<RibbonControl> InActivePygmentizeControls
        {
            get => new RibbonControl[] { lblPygmentsNotAvailable, btnHelpPygmentize };
        }
    }
}
