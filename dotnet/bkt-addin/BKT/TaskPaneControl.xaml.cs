/*
 * Created by SharpDevelop.
 * User: rdebeerst
 * 
 * WPF-UserControl for BKT-Content in the taskpane area.
 * Events are handled through the python-delegate
 *
 */

using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;

using System.Runtime.InteropServices;

using Fluent;

using System.Diagnostics; // for Debug



namespace BKT
{
	public partial class TaskPaneControl : UserControl
	{

        // link to the python delegate, where events are routed to
		private dynamic python_delegate;
        
        
		public TaskPaneControl()
		{
			InitializeComponent();
		}
		
        
        
        #region python delegate
        
        // set python delegate, where events can be reouted to
        public void SetPythonDelegate(dynamic new_python_delegate)
        {
            python_delegate = new_python_delegate;
        }
        
        // updares content of user control, according to configuration in Python-part
        public void UpdateContent()
        {
            try {
                FrameworkElement rootElement;
                
                // load xml from python
                DebugMessage("UpdateContent: obain taskpane xml-string");
                string xmlStr = python_delegate.get_custom_taskpane_ui();
                if (xmlStr == null)
                    return;
                // create root element
                DebugMessage("UpdateContent: parse xml, create markup, create layout");
                rootElement = (FrameworkElement) System.Windows.Markup.XamlReader.Parse(xmlStr);
                
                // add to layout
                layoutGrid.Children.Clear();
                layoutGrid.Children.Add(rootElement);
                
                
            } catch (Exception e) {
                //Debug.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss,fff") + ": " + e.ToString());
                DebugMessage(e.ToString());
                //System.Windows.Forms.MessageBox.Show(e.ToString());
            }   
        }
        
        #endregion
        
        
		
        #region events
        
        // handles any standard routed event
        private void Routed_Event(object sender, RoutedEventArgs e)
        {
            DebugMessage(String.Format("event TaskPane: {0} from {1} ", e.RoutedEvent, e.Source));
            //if (e.Handled == true) return;
            if (!python_delegate) return ;
            try {
                python_delegate.task_pane(sender, e);
                return ;
            } catch (Exception err) {
                Message(err.ToString());
                return ;
            }
        }
        
        #endregion
        
        
        
        #region Click events
        
        private void Click(object sender, RoutedEventArgs e)
        {
            DebugMessage(String.Format("event TaskPane: Click/{0} from {1} ", e.RoutedEvent, e.Source));
            //if (e.Handled == true) return;
            if (!python_delegate) return ;
            try {
                python_delegate.task_pane(sender, e); // FIXME: use on_click/on_action
                return ;
            } catch (Exception err) {
                Message(err.ToString());
                return ;
            }
        }
        
        private void Toggle_Click(object sender, RoutedEventArgs e)
        {
            DebugMessage(String.Format("event TaskPane: {0} from {1} ", e.RoutedEvent, e.Source));
            //if (e.Handled == true) return;
            if (!python_delegate) return ;
            try {
                python_delegate.task_pane(sender, e); // FIXME: use toggle_action with source.IsChecked
                return ;
            } catch (Exception err) {
                Message(err.ToString());
                return ;
            }
        }

        private void Menu_Click(object sender, RoutedEventArgs e)
        {
            DebugMessage(String.Format("event TaskPane: {0} from {1} ", e.RoutedEvent, e.Source));
            //if (e.Handled == true) return;
            if (!python_delegate) return ;
            try {
                if (( (System.Windows.Controls.MenuItem) e.Source).IsCheckable)
                {
                    Toggle_Click(sender, e);
                } else {
                    Click(sender, e);
                }
                return ;
            } catch (Exception err) {
                Message(err.ToString());
                return ;
            }
        }
                
        #endregion
        
        
        #region Spinner events
        
        private void Spinner_Loaded(object sender, RoutedEventArgs e)
        {
            DebugMessage(String.Format("event TaskPane: {0} from {1} ", e.RoutedEvent, e.Source));
            ((Fluent.Spinner)sender).ValueChanged += new RoutedPropertyChangedEventHandler<double>(Value_Changed_Event);
        }
        
        // handles Property changed event for Fluent.Spinner.ValueChanged
        private void Value_Changed_Event(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            DebugMessage(String.Format("event TaskPane: ValueChanged from {0} ", sender));
            if (!python_delegate) return ;
            try {
                python_delegate.task_pane_value_changed(sender, e);
                return ;
            } catch (Exception err) {
                Message(err.ToString());
                return ;
            }
        }
        
        #endregion
        
        
        #region TextBox events
                
        private void Text_LostFocus(object sender, RoutedEventArgs e)
        {
            DebugMessage(String.Format("event TaskPane: {0} from {1} ", e.RoutedEvent, e.Source));
            //if (e.Handled == true) return;
            if (!python_delegate) return ;
            try {
                python_delegate.task_pane(sender, e); // FIXME: use on_change
                return ;
            } catch (Exception err) {
                Message(err.ToString());
                return ;
            }
        }
        
        private new void KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            DebugMessage(String.Format("event TaskPane: {0} from {1} ", e.RoutedEvent, e.Source));
            if (!python_delegate) return ;
            try {
                if (e.Key == System.Windows.Input.Key.Return) {
                    // TODO: for ComboBox check IsReadOnly ?
                    python_delegate.task_pane(sender, e); // FIXME: use on_change, Source.Text
                }
                return ;
            } catch (Exception err) {
                Message(err.ToString());
                return ;
            }
        }
        
        #endregion
        
        
        #region ComboBox events
        
        private void Combo_LostFocus(object sender, RoutedEventArgs e)
        {
            DebugMessage(String.Format("event TaskPane: {0} from {1} ", e.RoutedEvent, e.Source));
            //if (e.Handled == true) return;
            if (!python_delegate) return ;
            try {
                if (! ( (Fluent.ComboBox) e.Source).IsReadOnly) {
                    python_delegate.task_pane(sender, e); // FIXME: use on_change
                    return ;
                }
                return ;
            } catch (Exception err) {
                Message(err.ToString());
                return ;
            }
        }
        
        private void Combo_SelectionChanged(object sender, RoutedEventArgs e)
        {
            DebugMessage(String.Format("event TaskPane: {0} from {1} ", e.RoutedEvent, e.Source));
            //if (e.Handled == true) return;
            if (!python_delegate) return ;
            try {
                python_delegate.task_pane(sender, e); // FIXME: use on_change and on_action_indexed
                return ;
            } catch (Exception err) {
                Message(err.ToString());
                return ;
            }
        }
        
        #endregion
        
        
        #region Gallery events
        
        private void Gallery_SelectionChanged(object sender, RoutedEventArgs e)
        {
            DebugMessage(String.Format("event TaskPane: {0} from {1} ", e.RoutedEvent, e.Source));
            //if (e.Handled == true) return;
            if (!python_delegate) return ;
            try {
                python_delegate.task_pane(sender, e); // FIXME: use on_change
                return ;
            } catch (Exception err) {
                Message(err.ToString());
                return ;
            }
        }
        
        private void SelectedColorChanged(object sender, RoutedEventArgs e)
        {
            DebugMessage(String.Format("event TaskPane: {0} from {1} ", e.RoutedEvent, e.Source));
            //if (e.Handled == true) return;
            if (!python_delegate) return ;
            try {
                python_delegate.task_pane(sender, e); // FIXME: use on_rgb_color_change
                return ;
            } catch (Exception err) {
                Message(err.ToString());
                return ;
            }
        }
            
        #endregion
        
        
        
        #region Date picker events
        
        private void SelectedDateChanged(object sender, RoutedEventArgs e)
        {
            DebugMessage(String.Format("event TaskPane: {0} from {1} ", e.RoutedEvent, e.Source));
            //if (e.Handled == true) return;
            if (!python_delegate) return ;
            try {
                python_delegate.task_pane(sender, e); // FIXME: use on_change
                return ;
            } catch (Exception err) {
                Message(err.ToString());
                return ;
            }
        }

        #endregion
        
        
        
        #region Debug
        
		private void Message(string s) {
			System.Windows.Forms.MessageBox.Show(s);
		}
        
        private void DebugMessage(string message)
        {
            Debug.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss,fff") + ": " + message);
        }
        
        #endregion
        
	}
}