using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Refractor
{
    public partial class Refractor : Form
    {
        public Refractor()
        {
            InitializeComponent();
            textBox5.Text = "C:\\Users\\" + Environment.UserName + "\\Downloads";
        }

        private void Get_Click(object sender, EventArgs e)
        {

            treeView1.Nodes.Clear();

            Assembly assem1 = Assembly.Load(AssemblyName.GetAssemblyName(textBox1.Text));

            Type[] types = assem1.GetTypes();
            int p = 0;
            int n = 0;
            int m = 0;
            int h = 0;

            foreach (Type tc in types)
            {
                TreeNode tnc = new TreeNode();

                if (tc.IsAbstract)
                {
                    tnc.Text = tc.Name;
                    tnc.ForeColor = Color.Red;
                }
                else if (tc.IsPublic)
                {
                    tnc.Text = tc.Name;
                    tnc.ForeColor = Color.Blue;
                    p++;
                }
                else if (tc.IsSealed)
                {
                    tnc.Text = tc.Name;
                    tnc.ForeColor = Color.DarkViolet;
                }
                else if(tc.IsInterface)
                {
                    tnc.Text = tc.Name;
                    tnc.ForeColor = Color.Blue;
                    tnc.BackColor = Color.Snow;
                }
                else if(tc.IsNested)
                {
                    tnc.Text = "{"+tc.Name+"}";
                    tnc.ForeColor = Color.Blue;
                    n++;
                }
                else if(tc.IsNotPublic)
                {
                    tnc.Text = tc.Name;
                    tnc.ForeColor = Color.Green;
                }
                else if(tc.IsAutoClass)
                {
                    tnc.Text = tc.Name;
                    tnc.ForeColor = Color.DarkKhaki;
                }
                else
                {
                    tnc.Text = tc.Name;
                    tnc.ForeColor = Color.DarkGray;
                }

                //Get List of Method Names of Class

                MemberInfo[] methodName = tc.GetMethods();
                MemberInfo[] propName = tc.GetProperties();

                List<MethodInfo> methods = ((System.Reflection.TypeInfo)tc).DeclaredMethods.ToList();
                List<PropertyInfo> props = ((System.Reflection.TypeInfo)tc).DeclaredProperties.ToList();
                //h += Convert.ToInt32(Math.Ceiling((methods.Count / 2) * 1.3));
                h += Convert.ToInt32((methods.Count));


                foreach (MemberInfo method in methodName)
                {
                    TreeNode tnm = new TreeNode();
                    if (method.Name.ToLower() == "tostring" || method.Name.ToLower() == "gethashcode" || method.Name.ToLower() == "equals" || method.Name.ToLower() == "gettype")
                    {
                        if (method.ReflectedType.IsPublic)
                        {
                            tnm.Text = method.Name;
                            tnm.ForeColor = Color.DarkOrange;
                        }
                        else
                        {
                            tnm.Text = method.Name;
                            tnm.ForeColor = Color.DarkRed;
                        }
                    }
                    else
                    {
                        if (method.ReflectedType.IsPublic)
                        {
                            tnm.Text = method.Name;
                            tnm.ForeColor = Color.DarkGreen;
                        }
                        else
                        {
                            tnm.Text = method.Name;
                            tnm.ForeColor = Color.Green;
                        }
                        m++;
                    }
                   
                    tnc.Nodes.Add(tnm);
                }
                foreach (MemberInfo prop in propName)
                {
                    TreeNode tnp = new TreeNode();
                    if (prop.ReflectedType.IsPublic)
                    {
                        tnp.Text = prop.Name;
                        tnp.ForeColor = Color.DarkSeaGreen;
                        tnp.BackColor = Color.Snow;
                    }
                    else
                    {
                        tnp.Text = prop.Name;
                        tnp.ForeColor = Color.DarkOrange;
                        tnp.BackColor = Color.Snow;
                    }

                    tnc.Nodes.Add(tnp);
                }
                treeView1.Nodes.Add(tnc);
            }
            textBox2.Text = (p+n).ToString();
            textBox3.Text = m.ToString();
            textBox4.Text = h.ToString();
        }

        private void fileSelect(object sender, MouseEventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog();
            openFileDialog1.InitialDirectory = @"C:\";
            openFileDialog1.Title = "Browse Dynamic Link Library(.dll) Files";
            openFileDialog1.Filter = "msi files (*.msi)|*.msi|exe files (*.exe)|*.exe|dll files (*.dll)|*.dll|All files (*.*)|*.*";
            textBox1.Text = openFileDialog1.FileName;
            openFileDialog1.Multiselect = true;
        }

        private void folderDialog(object sender, EventArgs e)
        {
            FolderBrowserDialog fd = new FolderBrowserDialog();
            fd.ShowDialog();
            textBox5.Text = fd.SelectedPath;
            fd.ShowNewFolderButton = true;
        }

        private void saveXL(object sender, EventArgs e)
        {
            DataSet ds = new DataSet();

            DataTable Classes = new DataTable();
            Classes.Columns.Add("Name");
            Classes.Columns.Add("Access");
            Classes.Columns.Add("Type");
            Classes.Columns.Add("FullName");
            Classes.Columns.Add("NameSpace");
            Classes.Columns.Add("Assembly");
            Classes.Columns.Add("Attributes");
            Classes.Columns.Add("BaseType");


            DataTable Methods = new DataTable();

            Methods.Columns.Add("Name");
            Methods.Columns.Add("Access");
            Methods.Columns.Add("ClassName");
            Methods.Columns.Add("Parameters");
            Methods.Columns.Add("ReturnType");

            DataTable Properties = new DataTable();

            Properties.Columns.Add("Name");
            Properties.Columns.Add("ClassName");



            Assembly assem1 = Assembly.Load(AssemblyName.GetAssemblyName(textBox1.Text));

            Type[] types = assem1.GetTypes();

            foreach (Type tc in types)
            {
                if (tc.Name.ToCharArray()[0] == '<' || tc.Name.ToCharArray()[0] == '_') { }
                else
                {
                    if (tc.IsClass || tc.IsInterface)
                    {
                        string access = string.Empty; string type = string.Empty;
                        if (tc.IsInterface)
                        {
                            type = "Interface";
                        }
                        else
                        {
                            if (tc.IsAbstract)
                            {
                                type = "Abstract Class";
                            }
                            else if (tc.IsSealed)
                            {
                                type = "Sealed Class";
                            }
                            else if (tc.IsNested)
                            {
                                type = "{Nested} Class";
                            }
                            else if (tc.IsAutoClass)
                            {
                                type = "Auto Class";
                            }
                            else
                            {
                                type = "General/Concrete Class";
                            }
                        }


                        if (tc.IsPublic)
                        {
                            access = "Public";
                        }
                        else if (tc.IsNotPublic)
                        {
                            access = "Private or Protected or Internal or Default";
                        }
                        else
                        {
                            access = "Something Else ";
                        }

                        Classes.Rows.Add(tc.Name, access, type, tc.FullName, tc.Namespace, tc.AssemblyQualifiedName, tc.Attributes.ToString(), tc.BaseType);

                        MemberInfo[] methods = tc.GetMethods();
                        MemberInfo[] properties = tc.GetProperties();

                        List<MethodInfo> methodsList = ((System.Reflection.TypeInfo)tc).DeclaredMethods.ToList();
                        List<PropertyInfo> propertiesList = ((System.Reflection.TypeInfo)tc).DeclaredProperties.ToList();



                        foreach (MethodInfo method in methodsList)
                        {
                            if (method.Name.ToLower() == "tostring" || method.Name.ToLower() == "gethashcode" || method.Name.ToLower() == "equals" || method.Name.ToLower() == "gettype")
                            {

                            }
                            else
                            {
                                List<ParameterInfo> paramss = method.GetParameters().ToList();
                                List<string> paras = new List<string>();
                                if (paramss.Count > 0)
                                    paramss.ForEach(p => paras.Add(p.ParameterType + "  " + p.Name));

                                

                                string accessM = string.Empty;

                                if (method.IsPrivate)
                                {
                                    accessM = "Private";
                                }
                                else if (method.IsPublic)
                                {
                                    accessM = "Public";
                                }
                                else
                                {
                                    accessM = "Protected or Internal";
                                }
                                

                                Methods.Rows.Add(method.Name, accessM, tc.Name, string.Join(" , ", paras), method.ReturnType);
                            }
                        }
                        foreach (PropertyInfo property in propertiesList)
                        {
                            Properties.Rows.Add(property.Name, tc.Name);
                        }
                    }
                }
            }
            ds.Tables.Add(Classes);
            ds.Tables.Add(Methods);
            ds.Tables.Add(Properties);

            string f = textBox1.Text;
            string[] ff = f.Split('\\');

            string filename = ff[ff.Length - 1].Split('.')[0] + ".xlsx";

            new ExcelOperations().ExportDataSet(ds, Path.Combine(textBox5.Text, filename));
            //new ExcelOperations().ExportDataSet(ds, textBox5.Text);
        }
    }
}

