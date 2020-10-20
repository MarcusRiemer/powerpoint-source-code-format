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

        }

        private void OnRenameSingle(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var selectedShapes = ActiveWindow.Selection.ShapeRange;
                foreach (Shape shape in selectedShapes)
                {
                    Formatter.FormatShape(shape);
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message, ex.GetType().Name);
            }
        }

        private Application Application
        {
            get => Globals.ThisAddIn.Application;
        }

        private Presentation ActivePresentation
        {
            get => Application.ActivePresentation;
        }

        private DocumentWindow ActiveWindow
        {
            get => Application.ActiveWindow;
        }
    }
}
