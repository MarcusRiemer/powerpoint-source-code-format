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
        private void RibbonLoad(object sender, RibbonUIEventArgs e)
        {
            cmbLanguage.Text = CurrentSettings.SelectedLanguage;
            cmbStyle.Text = CurrentSettings.SelectedStyle;
            ReflectPygmentizeStatus();
        }

        /// <summary>
        /// As it is not exactly straightforwad to get hold of pygmentize the UI
        /// tries to give at least a little feedback.
        /// </summary>
        private void ReflectPygmentizeStatus()
        {
            if (Pygments.FoundPygmentize)
            {
                lblPygmentsAvailable.SuperTip = Pygments.PygmentizePath;

                SetBoxVisible(bxAvailable, true);
                SetBoxVisible(bxUnavailable, false);
            }
            else
            {
                SetBoxVisible(bxAvailable, false);
                SetBoxVisible(bxUnavailable, true);
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

        private static void SetBoxEnabled(RibbonBox box, bool enabled)
        {
            box.Enabled = enabled;
            foreach (var item in box.Items)
            {
                item.Enabled = enabled;
            }
        }

        private static void SetBoxVisible(RibbonBox box, bool visible)
        {
            box.Visible = visible;
            foreach (var item in box.Items)
            {
                item.Visible = visible;
            }
        }

        /// <summary>
        /// The user has decided to use a different language, we want to remember this
        /// </summary>
        private void OnLanguageChanged(object sender, RibbonControlEventArgs e)
        {
            CurrentSettings.SelectedLanguage = cmbLanguage.Text;
            CurrentSettings.Save();
        }

        /// <summary>
        /// The user has decided to use a different style, we want to remember this
        /// </summary>
        private void OnStyleChanged(object sender, RibbonControlEventArgs e)
        {
            CurrentSettings.SelectedStyle = cmbStyle.Text;
            CurrentSettings.Save();
        }

        /// <summary>
        /// The user has some trouble setting up pygments, lets show him where he can get help
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnShowOnlineHelp(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start("https://github.com/MarcusRiemer/powerpoint-source-code-format");
        }

        /// <summary>
        /// The settings that should be applicable to the current user
        /// </summary>
        private Properties.Settings CurrentSettings {
            get => Properties.Settings.Default;
        }
    }
}
