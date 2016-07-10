using System;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Xml;

namespace SunnyExcel
{
    public class CLayout
    {
        private static CConfig __cconfig = new CConfig();

        //-------------------------------------------------------------------------------------------------------------------------
        //
        //-------------------------------------------------------------------------------------------------------------------------
        private string GetLayoutFileName(string p_form_id, string p_name)
        {
            string _folder = Path.Combine(__cconfig.GetWorkingFolder("LayoutHelper"), p_form_id);

            if (Directory.Exists(_folder) == false)
                Directory.CreateDirectory(_folder);

            return Path.Combine(_folder, String.Format("{0}.xml", p_name));
        }

        private DataTable LayoutDataTable()
        {
            var _dt = new DataTable("LayoutSaver");

            var _keys = new DataColumn[1];
            {
                _keys[0] = new DataColumn("name", typeof(String));
                _dt.Columns.Add(_keys[0]);

                _dt.Columns.Add("left", typeof(Int32));
                _dt.Columns.Add("top", typeof(Int32));
                _dt.Columns.Add("width", typeof(Int32));
                _dt.Columns.Add("height", typeof(Int32));

                _dt.Columns.Add("state", typeof(Int32));
                _dt.Columns.Add("reset", typeof(Boolean));

                _dt.Columns.Add("version", typeof(String));
                _dt.Columns.Add("text", typeof(String));
                _dt.Columns.Add("value", typeof(Object));

                _dt.Columns.Add("horizontal", typeof(Int32));
                _dt.Columns.Add("visibility", typeof(Int32));

                _dt.PrimaryKey = _keys;
            }

            return _dt;
        }

        private void SaveControlLayout(System.Windows.Forms.Control.ControlCollection p_ctrls, string p_form_id, ref DataTable p_dt)
        {
            foreach (System.Windows.Forms.Control _ctrl in p_ctrls)
            {
                if (_ctrl.GetType().Name == "ToolStrip")
                {
                    var _tool_strip = (_ctrl as ToolStrip);
                    foreach (ToolStripItem _item in _tool_strip.Items)
                    {
                        if (_item.GetType().Name != "ToolStripComboBox")
                            continue;

                        var _field_name = String.Format("{0}_{1}_{2}", _tool_strip.Parent.Name, _tool_strip.Name, _item.Name);

                        var _dr = p_dt.NewRow();
                        {
                            _dr["name"] = _field_name;
                            _dr["text"] = _item.Text;

                            p_dt.Rows.Add(_dr);
                        }
                    }
                }
                else if (_ctrl.GetType().Name == "Splitter")
                {
                    var _splitter = (_ctrl as Splitter);
                    var _field_name = String.Format("{0}_{1}", _splitter.Parent.Name, _splitter.Name);

                    var _dr = p_dt.NewRow();
                    {
                        _dr["name"] = _field_name;
                        _dr["top"] = _splitter.SplitPosition;

                        p_dt.Rows.Add(_dr);
                    }
                }
                else if (_ctrl.GetType().Name == "SplitContainer")
                {
                    var _spctrl = (_ctrl as SplitContainer);
                    var _field_name = String.Format("{0}_{1}", _spctrl.Parent.Name, _spctrl.Name);

                    var _dr = p_dt.NewRow();
                    {
                        _dr["name"] = _field_name;
                        _dr["top"] = _spctrl.Location.Y;

                        p_dt.Rows.Add(_dr);
                    }
                }
                else if (_ctrl.GetType().Name == "CheckBox")
                {
                    var _chkbox = (_ctrl as CheckBox);
                    var _field_name = String.Format("{0}_{1}", _chkbox.Parent.Name, _chkbox.Name);

                    var _dr = p_dt.NewRow();
                    {
                        _dr["name"] = _field_name;
                        _dr["value"] = _chkbox.Checked;

                        p_dt.Rows.Add(_dr);
                    }
                }
                else if (_ctrl.GetType().Name == "TextBox")
                {
                    var _txtbox = (_ctrl as TextBox);

                    if (_txtbox.ReadOnly == false)
                    {
                        var _field_name = String.Format("{0}_{1}", _txtbox.Parent.Name, _txtbox.Name);

                        var _dr = p_dt.NewRow();
                        {
                            _dr["name"] = _field_name;
                            _dr["text"] = _txtbox.Text;

                            p_dt.Rows.Add(_dr);
                        }
                    }
                }

                if (_ctrl.Controls != null && _ctrl.HasChildren == true)
                    SaveControlLayout(_ctrl.Controls, p_form_id, ref p_dt);
            }
        }

        //-------------------------------------------------------------------------------------------------------------------------
        // XtraGrid Verify Functions
        //-------------------------------------------------------------------------------------------------------------------------
        private string GetChildValue(XmlNode p_node, string p_name, string p_attname, string p_attval)
        {
            var _result = "";

            p_node = p_node.FirstChild;
            while (p_node != null)
            {
                if (p_node.Name.IndexOf(p_name) != -1)
                {
                    if (p_node.Attributes[p_attname].Value == p_attval)
                    {
                        _result = p_node.FirstChild.Value;
                        break;
                    }
                }

                p_node = p_node.NextSibling;
            }

            return _result;
        }

        private XmlNode GetChildNode(XmlNode p_node, string p_name, string p_attname, string p_attval)
        {
            XmlNode _result = null;

            p_node = p_node.FirstChild;
            while (p_node != null)
            {
                if (p_node.Name.IndexOf(p_name) != -1)
                {
                    if (p_node.Attributes[p_attname].Value == p_attval)
                    {
                        _result = p_node;
                        break;
                    }
                }

                p_node = p_node.NextSibling;
            }

            return _result;
        }


        //-------------------------------------------------------------------------------------------------------------------------
        // RestoreControlLayout
        //-------------------------------------------------------------------------------------------------------------------------
        private void RestoreControlLayout(System.Windows.Forms.Control.ControlCollection p_ctrls, string p_form_id, ref DataTable p_dt)
        {
            foreach (System.Windows.Forms.Control _ctrl in p_ctrls)
            {
                if (_ctrl.GetType().Name == "ToolStrip")
                {
                    var _tool_strip = (_ctrl as ToolStrip);
                    foreach (ToolStripItem _item in _tool_strip.Items)
                    {
                        if (_item.GetType().Name != "ToolStripComboBox")
                            continue;

                        var _field_name = String.Format("{0}_{1}_{2}", _tool_strip.Parent.Name, _tool_strip.Name, _item.Name);
                        var _dr = p_dt.Rows.Find(_field_name);
                        if (_dr != null)
                        {
                            _item.Text = Convert.ToString(_dr["text"]);
                        }
                    }
                }
                else if (_ctrl.GetType().Name == "Splitter")
                {
                    var _splitter = (_ctrl as Splitter);
                    var _field_name = String.Format("{0}_{1}", _splitter.Parent.Name, _splitter.Name);

                    var _dr = p_dt.Rows.Find(_field_name);
                    if (_dr != null)
                    {
                        _splitter.SplitPosition = Convert.ToInt32(_dr["top"]);
                    }
                }
                else if (_ctrl.GetType().Name == "SplitContainer")
                {
                    var _spctrl = (_ctrl as SplitContainer);
                    var _field_name = String.Format("{0}_{1}", _spctrl.Parent.Name, _spctrl.Name);

                    var _dr = p_dt.Rows.Find(_field_name);
                    if (_dr != null)
                    {
                        //_spctrl.Location.Y = Convert.ToInt32(_dr["top"]);
                    }
                }
                else if (_ctrl.GetType().Name == "CheckBox")
                {
                    var _chkbox = (_ctrl as CheckBox);
                    var _field_name = String.Format("{0}_{1}", _chkbox.Parent.Name, _chkbox.Name);

                    var _dr = p_dt.Rows.Find(_field_name);
                    if (_dr != null)
                    {
                        _chkbox.Checked = Convert.ToBoolean(_dr["value"]);
                    }
                }
                else if (_ctrl.GetType().Name == "TextBox")
                {
                    var _txtbox = (_ctrl as TextBox);

                    if (_txtbox.ReadOnly == false)
                    {
                        var _field_name = String.Format("{0}_{1}", _txtbox.Parent.Name, _txtbox.Name);

                        var _dr = p_dt.Rows.Find(_field_name);
                        if (_dr != null)
                        {
                            _txtbox.Text = Convert.ToString(_dr["text"]);
                        }
                    }
                }

                if (_ctrl.Controls != null && _ctrl.HasChildren == true)
                    RestoreControlLayout(_ctrl.Controls, p_form_id, ref p_dt);
            }
        }

        private void RemoveControlLayout(System.Windows.Forms.Control.ControlCollection p_ctrls, string p_form_id)
        {
            foreach (System.Windows.Forms.Control _ctrl in p_ctrls)
            {
                if (_ctrl.GetType().Name == "DataGridView")
                {
                    var _pvgrid = (_ctrl as DataGridView);

                    var _filename = GetLayoutFileName(p_form_id, _pvgrid.Name);
                    File.Delete(_filename);
                }

                if (_ctrl.Controls != null && _ctrl.HasChildren == true)
                    RemoveControlLayout(_ctrl.Controls, p_form_id);
            }
        }

        //-------------------------------------------------------------------------------------------------------------------------
        // RemoveFormLayout
        //-------------------------------------------------------------------------------------------------------------------------

        /// <summary>
        /// 
        /// </summary>
        /// <param name="p_user_form"></param>
        public void RemoveFormLayout(Form p_user_form)
        {
            RemoveFormLayout(p_user_form, p_user_form.GetType().GUID.ToString());
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="p_user_form"></param>
        /// <param name="p_form_id"></param>
        public void RemoveFormLayout(Form p_user_form, string p_form_id)
        {
            try
            {
                RemoveControlLayout(p_user_form.Controls, p_form_id);

                var _dt = LayoutDataTable();
                {
                    var _dr = _dt.NewRow();
                    _dr["name"] = p_form_id;
                    _dr["reset"] = true;

                    _dt.Rows.Add(_dr);
                }

                var _filename = GetLayoutFileName(p_form_id, p_user_form.Name);
                _dt.WriteXml(_filename);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="p_user_control"></param>
        public void RemoveFormLayout(UserControl p_user_control)
        {
            RemoveFormLayout(p_user_control, p_user_control.GetType().GUID.ToString());
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="p_user_control"></param>
        /// <param name="p_form_id"></param>
        public void RemoveFormLayout(UserControl p_user_control, string p_form_id)
        {
            try
            {
                RemoveControlLayout(p_user_control.Controls, p_form_id);

                var _dt = LayoutDataTable();
                {
                    var _dr = _dt.NewRow();
                    _dr["name"] = p_form_id;
                    _dr["reset"] = true;

                    _dt.Rows.Add(_dr);
                }

                var _filename = GetLayoutFileName(p_form_id, p_user_control.Name);
                _dt.WriteXml(_filename);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
            }
        }

        //-------------------------------------------------------------------------------------------------------------------------
        // RestoreFormLayout
        //-------------------------------------------------------------------------------------------------------------------------

        /// <summary>
        /// 
        /// </summary>
        /// <param name="p_user_form"></param>
        public void RestoreFormLayout(Form p_user_form)
        {
            RestoreFormLayout(p_user_form, p_user_form.GetType().GUID.ToString());
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="p_user_form"></param>
        /// <param name="p_form_id"></param>
        public void RestoreFormLayout(Form p_user_form, string p_form_id)
        {
            try
            {
                if (p_user_form.WindowState != FormWindowState.Maximized)
                {
                    var _filename = GetLayoutFileName(p_form_id, p_user_form.Name);
                    if (File.Exists(_filename) == true)
                    {
                        var _dt = LayoutDataTable();
                        _dt.ReadXml(_filename);

                        var _read = true;

                        var _dr = _dt.Rows.Find(p_form_id);
                        if (_dr != null)
                        {
                            if (Convert.ToBoolean(_dr["reset"]) == true)
                                _read = false;
                        }

                        if (_read == true)
                        {
                            _dr = _dt.Rows.Find(p_user_form.Name);
                            if (_dr != null)
                            {
                                var _version = Convert.ToString(_dr["version"]);
                                if (_version == p_user_form.GetType().Assembly.GetName().Version.ToString())
                                {
                                    var _left = Convert.ToInt32(_dr["left"]);
                                    var _top = Convert.ToInt32(_dr["top"]);
                                    var _width = Convert.ToInt32(_dr["width"]);
                                    var _height = Convert.ToInt32(_dr["height"]);

                                    var _state = (FormWindowState)Convert.ToInt32(_dr["state"]);
                                    if (_state != FormWindowState.Maximized)
                                    {
                                        p_user_form.Location = new Point(_left, _top);
                                        p_user_form.Size = new Size(_width, _height);
                                    }

                                    p_user_form.WindowState = _state;

                                    RestoreControlLayout(p_user_form.Controls, p_form_id, ref _dt);
                                }
                            }
                        }
                    }
                }

                if (p_user_form.Site != null)
                {
                    foreach (System.ComponentModel.Component _c in p_user_form.Site.Container.Components)
                    {
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="p_user_control"></param>
        /// <param name="p_user_form"></param>
        public void RestoreFormLayout(UserControl p_user_control, Form p_user_form)
        {
            RestoreFormLayout(p_user_form, p_user_control.GetType().GUID.ToString());
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="p_user_control"></param>
        public void RestoreFormLayout(UserControl p_user_control)
        {
            RestoreFormLayout(p_user_control, p_user_control.GetType().GUID.ToString());
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="p_user_control"></param>
        /// <param name="p_form_id"></param>
        public void RestoreFormLayout(UserControl p_user_control, string p_form_id)
        {
            try
            {
                var _filename = GetLayoutFileName(p_form_id, p_user_control.Name);
                if (File.Exists(_filename) == true)
                {
                    var _dt = LayoutDataTable();
                    _dt.ReadXml(_filename);

                    var _read = true;

                    var _dr = _dt.Rows.Find(p_form_id);
                    if (_dr != null)
                    {
                        if (Convert.ToBoolean(_dr["reset"]) == true)
                            _read = false;
                    }

                    if (_read == true)
                    {
                        _dr = _dt.Rows.Find(p_user_control.Name);
                        if (_dr != null)
                        {
                            string _version = Convert.ToString(_dr["version"]);
                            if (_version == p_user_control.GetType().Assembly.GetName().Version.ToString())
                            {
                                RestoreControlLayout(p_user_control.Controls, p_form_id, ref _dt);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
            }
        }

        //-------------------------------------------------------------------------------------------------------------------------
        // SaveFormLayout
        //-------------------------------------------------------------------------------------------------------------------------

        /// <summary>
        /// 
        /// </summary>
        /// <param name="p_user_form"></param>
        public void SaveFormLayout(Form p_user_form)
        {
            SaveFormLayout(p_user_form, p_user_form.GetType().GUID.ToString());
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="p_user_form"></param>
        /// <param name="p_form_id"></param>
        public void SaveFormLayout(Form p_user_form, string p_form_id)
        {
            try
            {
                var _write = true;

                if (p_user_form.Site != null)
                {
                    foreach (System.ComponentModel.Component _c in p_user_form.Site.Container.Components)
                    {
                    }
                }

                var _filename = GetLayoutFileName(p_form_id, p_user_form.Name);
                if (File.Exists(_filename) == true)
                {
                    var _dt = LayoutDataTable();
                    _dt.ReadXml(_filename);

                    var _dr = _dt.Rows.Find(p_form_id);
                    if (_dr != null && Convert.ToBoolean(_dr["reset"]) == true)
                    {
                        File.Delete(_filename);
                        _write = false;
                    }
                }

                if (_write == true)
                {
                    DataTable _dt = LayoutDataTable();
                    {
                        var _dr = _dt.NewRow();
                        _dr["name"] = p_user_form.Name;
                        _dr["version"] = p_user_form.GetType().Assembly.GetName().Version.ToString();

                        var _state = p_user_form.WindowState;
                        if (_state == FormWindowState.Minimized)
                            _state = FormWindowState.Normal;
                        else
                        {
                            _dr["left"] = p_user_form.Location.X;
                            _dr["top"] = p_user_form.Location.Y;
                            _dr["width"] = p_user_form.Size.Width;
                            _dr["height"] = p_user_form.Size.Height;
                        }

                        _dr["state"] = (int)_state;

                        _dt.Rows.Add(_dr);
                    }

                    SaveControlLayout(p_user_form.Controls, p_form_id, ref _dt);

                    _dt.WriteXml(_filename);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="p_user_control"></param>
        /// <param name="p_user_form"></param>
        public void SaveFormLayout(UserControl p_user_control, Form p_user_form)
        {
            SaveFormLayout(p_user_form, p_user_control.GetType().GUID.ToString());
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="p_user_control"></param>
        public void SaveFormLayout(UserControl p_user_control)
        {
            SaveFormLayout(p_user_control, p_user_control.GetType().GUID.ToString());
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="p_user_control"></param>
        /// <param name="p_form_id"></param>
        public void SaveFormLayout(UserControl p_user_control, string p_form_id)
        {
            try
            {
                bool _write = true;

                string _filename = GetLayoutFileName(p_form_id, p_user_control.Name);
                if (File.Exists(_filename) == true)
                {
                    DataTable _dt = LayoutDataTable();
                    _dt.ReadXml(_filename);

                    DataRow _dr = _dt.Rows.Find(p_form_id);
                    if (_dr != null && Convert.ToBoolean(_dr["reset"]) == true)
                    {
                        File.Delete(_filename);
                        _write = false;
                    }
                }

                if (_write == true)
                {
                    DataTable _dt = LayoutDataTable();
                    {
                        DataRow _dr = _dt.NewRow();
                        _dr["name"] = p_user_control.Name;
                        _dr["version"] = p_user_control.GetType().Assembly.GetName().Version.ToString();

                        _dt.Rows.Add(_dr);
                    }

                    SaveControlLayout(p_user_control.Controls, p_form_id, ref _dt);

                    _dt.WriteXml(_filename);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
            }
        }

        //-------------------------------------------------------------------------------------------------------------------------
        //
        //-------------------------------------------------------------------------------------------------------------------------
    }
}