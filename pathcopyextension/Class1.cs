using System;

using SharpShell.Attributes;
using SharpShell.SharpContextMenu;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.Diagnostics;
using System.IO;

namespace pathcopyextension
{
    [ComVisible(true)]
    [Guid("210A97C8-A4ED-47E9-825E-7C8FD0CA084C")]
    [COMServerAssociation(AssociationType.AllFilesAndFolders)]
    //他のSharpShellを使用するソフトウェアとclass名が重複していると正常に機能しない
    public class NekotadonPathCopyExtension : SharpContextMenu
    {
        private ContextMenuStrip menu = null;
        List<menu_cmd> menu_cmd_lists = new List<menu_cmd>();

        private class menu_cmd
        {
            public long idx;
            public string cmd;
            public string arg;
            public bool each;

            public menu_cmd(long i, string c, string a, bool e)
            {
                idx = i;
                cmd = c;
                arg = a;
                each = e;
            }
        }
        private bool doublequotationmark = false;
        private bool lastNewLine = false;

        protected override bool CanShowMenu()
        {
            try
            {
                //setting.xmlのパス
                string folder = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                string file = folder + @"\setting.xml";

                //copy pathを表示するかどうか
                bool copypath = true;
                if (System.IO.File.Exists(file))
                {
                    //xmlファイルを指定する
                    XElement xml = XElement.Load(file);

                    //メンバー情報のタグ内の情報を取得する
                    IEnumerable<XElement> cpinfos = from item in xml.Elements("copypath") select item;

                    if (cpinfos.Count() == 1)
                    {
                        foreach (XElement ele in cpinfos.Elements())
                        {
                            if (ele.Name == "enable")
                            {
                                if (ele.Value == "false")
                                {
                                    copypath = false;
                                }
                            }
                            else if (ele.Name == "doublequotationmark")
                            {
                                if (ele.Value == "true")
                                {
                                    doublequotationmark = true;
                                }
                            }
                            else if (ele.Name == "lastNewLine")
                            {
                                if (ele.Value == "true")
                                {
                                    lastNewLine = true;
                                }
                            }
                        }
                    }
                }

                if (copypath)
                {
                    menu = new ContextMenuStrip();

                    var itemcopypath = new ToolStripMenuItem
                    {
                        Text = "Copy Path",
                    };

                    itemcopypath.Click += new System.EventHandler(this.copypathaction);

                    menu.Items.Add(itemcopypath);
                }

                //ファイルごとのメニュー

                bool is_exist_dir = false;
                bool is_exist_file = false;
                bool one_ext = true;
                string kakutyoushi = "";
                foreach (var filePath in SelectedItemPaths)
                {
                    if (System.IO.Directory.Exists(filePath))
                    {
                        //ディレクトリがある場合は終了
                        is_exist_dir = true;
                        break;
                    }
                    else if (System.IO.File.Exists(filePath))
                    {
                        if (System.IO.Path.GetExtension(filePath) != "")
                        {
                            //ファイルが存在して拡張子がある場合
                            is_exist_file = true;
                            if (kakutyoushi == "")
                            {
                                kakutyoushi = System.IO.Path.GetExtension(filePath).ToLower();
                            }
                            else
                            {
                                if (kakutyoushi != System.IO.Path.GetExtension(filePath).ToLower())
                                {
                                    //複数拡張子がある場合は終了
                                    one_ext = false;
                                    break;
                                }
                            }
                        }
                    }
                }

                if (System.IO.File.Exists(file) && !is_exist_dir && is_exist_file && one_ext && kakutyoushi != "")
                {
                    //xmlファイルを指定する
                    XElement xml = XElement.Load(file);

                    //メンバー情報のタグ内の情報を取得する
                    IEnumerable<XElement> extmenus = from item in xml.Elements("extmenu") select item;

                    foreach (XElement extmenu in extmenus)
                    {
                        var ext = extmenu.Attribute("ext");
                        if (ext != null && ext.Value.ToLower() == kakutyoushi)
                        {
                            if (menu == null)
                            {
                                menu = new ContextMenuStrip();
                            }

                            if (extmenu.HasElements)
                            {
                                foreach (XElement child in extmenu.Elements())
                                {
                                    if (child.Name == "button")
                                    {
                                        ToolStripMenuItem button_item = button_to_menu(child);

                                        if (button_item != null)
                                        {
                                            menu.Items.Add(button_item);
                                        }
                                    }
                                    else if (child.Name == "menu")
                                    {
                                        ToolStripMenuItem menu_item = menu_to_menu(child);

                                        if (menu_item != null)
                                        {
                                            menu.Items.Add(menu_item);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                ;
            }

            return menu != null;
        }

        protected override ContextMenuStrip CreateMenu()
        {
            if (menu == null)
            {
                menu = new ContextMenuStrip();
            }

            return menu;
        }

        private void copypathaction(object sender, EventArgs e)
        {
            StringBuilder sb = new StringBuilder();
            string crlf = Environment.NewLine;
            bool first = true;

            foreach (var filePath in SelectedItemPaths)
            {
                if (first)
                {
                    if (doublequotationmark)
                    {
                        sb.Append("\"");
                    }
                    sb.Append(filePath);
                    if (doublequotationmark)
                    {
                        sb.Append("\"");
                    }
                    first = false;
                }
                else
                {
                    sb.Append(crlf);
                    if (doublequotationmark)
                    {
                        sb.Append("\"");
                    }
                    sb.Append(filePath);
                    if (doublequotationmark)
                    {
                        sb.Append("\"");
                    }
                }
            }
            if (lastNewLine)
            {
                sb.Append(crlf);
            }

            Clipboard.SetText(sb.ToString());
        }

        private ToolStripMenuItem button_to_menu(XElement button)
        {
            ToolStripMenuItem ret = null;

            try
            {
                if (button.HasElements && button.Elements().Count() == 4)
                {
                    int c_name = 0;
                    int c_command = 0;
                    int c_arg = 0;
                    int c_each = 0;

                    foreach (XElement ele in button.Elements())
                    {
                        if (ele.Name == "name")
                        {
                            c_name++;
                        }
                        else if (ele.Name == "command")
                        {
                            c_command++;
                        }
                        else if (ele.Name == "arg")
                        {
                            c_arg++;
                        }
                        else if (ele.Name == "each")
                        {
                            c_each++;
                        }
                    }

                    if (c_name == 1 && c_command == 1 && c_arg == 1 && c_each == 1)
                    {
                        ret = new ToolStripMenuItem();
                        ret.Text = button.Element("name").Value;
                        ret.Name = "cmd" + menu_cmd_lists.Count.ToString();
                        ret.Click += new System.EventHandler(this.click_func);
                        menu_cmd_lists.Add(new menu_cmd(menu_cmd_lists.Count, button.Element("command").Value, button.Element("arg").Value, button.Element("each").Value == "true" ? true : false));
                        return ret;
                    }
                }
            }
            catch (Exception)
            {
                ;
            }

            return ret;
        }
        private ToolStripMenuItem menu_to_menu(XElement menu)
        {
            ToolStripMenuItem ret = null;

            try
            {
                var menu_name = menu.Attribute("name");

                if (menu_name != null)
                {
                    ret = new ToolStripMenuItem();
                    ret.Text = menu_name.Value;
                    foreach (XElement child in menu.Elements())
                    {
                        if (child.Name == "button")
                        {
                            ToolStripMenuItem button_item = button_to_menu(child);

                            if (button_item != null)
                            {
                                ret.DropDownItems.Add(button_item);
                            }
                        }
                        else if (child.Name == "menu")
                        {
                            ToolStripMenuItem menu_item = menu_to_menu(child);

                            if (menu_item != null)
                            {
                                ret.DropDownItems.Add(menu_item);
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                ;
            }

            return ret;
        }
        private string filename_to_filepath(string file)
        {
            string folderpath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string folderpath_yen = folderpath + @"\";

            string f = file;

            try
            {
                if (f.StartsWith(@"..\"))//相対パスの場合
                {
                    f = Path.GetDirectoryName(folderpath) + f.Substring(2);
                }
                else if (f.StartsWith(@".\"))//相対パスの場合
                {
                    f = folderpath_yen + f.Substring(2);
                }
                else if (!f.Contains(@"\") && File.Exists(folderpath_yen + f))//直下のファイルの場合
                {
                    f = folderpath_yen + f;
                }
            }
            catch (Exception)
            {
                ;
            }

            return f;
        }
        private void click_func(object sender, EventArgs e)
        {
            long n = 0;
            if (typeof(ToolStripMenuItem) == sender.GetType())
            {
                n = long.Parse(((ToolStripMenuItem)sender).Name.Replace("cmd", ""));
            }

            foreach (var s in menu_cmd_lists)
            {
                if (s.idx == n)
                {
                    if (s.each)
                    {
                        //個々に処理
                        foreach (var filePath in SelectedItemPaths)
                        {
                            string file = filePath;

                            if (file.Contains(" "))
                            {
                                file = "\"" + file + "\"";
                            }

                            string arg = s.arg.Replace("__file__", file);

                            if (System.IO.File.Exists(filePath))
                            {
                                try
                                {
                                    Process p = new Process();
                                    p.StartInfo.FileName = filename_to_filepath(s.cmd);
                                    p.StartInfo.Arguments = arg;
                                    p.StartInfo.UseShellExecute = false;
                                    p.StartInfo.WorkingDirectory = System.IO.Path.GetDirectoryName(filePath);
                                    p.Start();
                                }
                                catch (Exception)
                                {

                                }
                            }
                        }
                    }
                    else
                    {
                        //まとめて処理
                        string files = "";
                        bool exist = true;
                        string lastfile = "";
                        foreach (var filePath in SelectedItemPaths)
                        {
                            lastfile = filePath;
                            if (filePath.Contains(" "))
                            {
                                files += "\"" + filePath + "\"" + " ";
                            }
                            else
                            {
                                files += filePath + " ";
                            }

                            if (!System.IO.File.Exists(filePath))
                            {
                                exist = false;
                            }
                        }

                        if (exist)
                        {
                            string arg = s.arg.Replace("__file__", files);

                            try
                            {
                                Process p = new Process();
                                p.StartInfo.FileName = filename_to_filepath(s.cmd);
                                p.StartInfo.Arguments = arg;
                                p.StartInfo.UseShellExecute = false;
                                p.StartInfo.WorkingDirectory = System.IO.Path.GetDirectoryName(lastfile);

                                p.Start();
                            }
                            catch (Exception)
                            {

                            }
                        }
                    }
                    break;
                }
            }
        }
    }
}


