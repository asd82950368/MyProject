using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop.Excel;
namespace homework
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            comboBox1.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox1.SelectedIndex = 0;
            textBox1.Multiline = false;
            textBox2.Multiline = false;
            textBox3.Multiline = false;
            NameMsgLab.Text = "";
            NumberMsgLab.Text = "";
            CodeMsgLab.Text = "";
        }

        private void textBox3_TabIndexChanged(object sender, EventArgs e)
        {
           
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            CheckAll cka = new CheckAll();
            ReadAndWrite red = new ReadAndWrite();

            string Name = textBox1.Text;            
            string NameMsg = cka.CheckName(Name);
            string number = textBox2.Text;           
            string NumberMsg = cka.CheckNumber(number);
            string code = textBox3.Text;            
            string CodeMsg = cka.CheckCode(code);
            string sex = comboBox1.Text;

            if (NameMsg is "OK" && NumberMsg is "OK" && CodeMsg is "OK")
            {
                NameMsgLab.Text = "";
                NumberMsgLab.Text = "";
                CodeMsgLab.Text = "";
                string codeGet = cka.CodeGet(sex, System.Convert.ToInt32(code));
                string OkMsg = "姓名:  " + Name + "\n學號:  "+ number + "\n性別:  "+ sex + "\n分數:  "+ code + "\n評語:  " + codeGet;
                string Data = Name +","+ number+"," + sex + "," + code + "," + codeGet;
                int coun = treeView1.GetNodeCount(false);
                treeView1.Nodes.Add("姓名:  " + Name);
                treeView1.Nodes[coun].Nodes.Add("學號:  " + number);
                treeView1.Nodes[coun].Nodes.Add("性別:  " + sex);
                treeView1.Nodes[coun].Nodes.Add("分數:  " + code);
                treeView1.Nodes[coun].Nodes.Add("評語:  " + codeGet);
                //string WritwOk = red.Write(Data);
                
               // MessageBox.Show(WritwOk+""+ OkMsg, "分數已送出");
            }
            else
            {
                NameMsgLab.Text = cka.CheckName(Name);
                NumberMsgLab.Text = cka.CheckNumber(number);
                CodeMsgLab.Text = cka.CheckCode(code);
                
            }
        }
   
        class ReadAndWrite
        {
            private string FilePath;
            private string msg = "資料寫入完成\n";
            
            public void Write(string data)
            {
                // 將字串寫入TXT檔
                StreamWriter str = new StreamWriter(@"E:\pixnet\20160614\Lab2_TXT_Read_Write\Write.TXT");
                string WriteWord = "aaaaa";
                str.WriteLine(WriteWord);
            
                str.Close();
            }
        }


        class CheckAll
        {
            private string CodeGetMsg;
            public string CodeGet(string SexGet ,int codeGet)
            {
                if (SexGet is "男")
                {
                    if (codeGet == 100)
                    {
                        CodeGetMsg = "真是太厲害了考了"+ codeGet + "分";
                    }
                    else if(codeGet >= 90)
                    {
                        CodeGetMsg = "考了"+ codeGet + "分也算厲害";
                    }
                    else if (codeGet >= 80)
                    {
                        CodeGetMsg = "考了"+ codeGet + "分還可以";
                    }
                    else if (codeGet >= 70)
                    {
                        CodeGetMsg = "考了" + codeGet + "分須再加強";
                    }
                    else if (codeGet >= 60)
                    {
                        CodeGetMsg = "考了" + codeGet + "分差勁了";
                    }
                    else
                    {
                        CodeGetMsg = "被當了謝謝";
                    }

                }
                else if(SexGet is "女")
                {
                    if (codeGet == 100)
                    {
                        CodeGetMsg = "真是太厲害了考了"+ codeGet + "分";
                    }
                    else if (codeGet >= 90)
                    {
                        CodeGetMsg = "考了" + codeGet + "分真強";
                    }
                    else if (codeGet >= 80)
                    {
                        
                        CodeGetMsg = "考了" + codeGet + "分也算厲害";
                    }
                    else if (codeGet >= 70)
                    {
                        
                        CodeGetMsg = "考了" + codeGet + "分還可以";
                    }
                    else if (codeGet >= 60)
                    {
                        
                        CodeGetMsg = "考了" + codeGet + "分須再加強";
                    }
                    else
                    {
                        CodeGetMsg = "被當了謝謝";
                    }
                }
                else
                {
                    CodeGetMsg = "特殊標準無法顯示";
                }
                return CodeGetMsg;
            }
            private string NameMsg;          
            public string CheckName(string Name)
            {
               
                if (Name != "")
                {
                    var chars = Name.ToCharArray();
                    for (int i = 0; i<chars.Length; i++)
                    {
                        char c = System.Convert.ToChar(chars[i]);
                        if (char.IsSeparator(c) == false && char.IsPunctuation(c) == false && char.IsSymbol(c) == false)
                        {
                            //MessageBox.Show("OK" + c);
                            NameMsg = "OK";
                        }
                        else
                        {
                            NameMsg = "輸入錯誤";
                            break;
                        }
                    }
                    
                }
                else
                {
                    NameMsg = "##您必須輸入"; 
                }
                return NameMsg;
            }
            private string NumberMsg;
            public string CheckNumber(string Number)
            {

                if (Number != "")
                {
                    
                    var chars = Number.ToCharArray();   
                    if (chars.Length == 9)
                    {
                        for (int i = 0; i < chars.Length; i++)
                        {
                            char c = System.Convert.ToChar(chars[i]);
                            //MessageBox.Show(System.Convert.ToString(chars[i]));
                            if (i <= 4 && char.IsNumber(c) == true)
                            {
                                
                            }
                            else if (i == 5 && char.IsLetter(c) == true)
                            {
                                
                            }
                            else if (i > 5  && char.IsNumber(c) == true)
                            {
                                NumberMsg = "OK";
                            }
                            else
                            {
                                NumberMsg = "##學號輸入錯誤";
                                break;                               
                            }                            
                        }
                    }
                    else
                    {
                        NumberMsg = "##學號為9個字元";
                    }                                     
                }
                else
                {
                    NumberMsg = "##您必須輸入";
                }
                return NumberMsg;
            }
            private string CodeMsg;
            public string CheckCode(string Code)
            {
                var code = Code.ToCharArray();
                if (Code != "")
                {                 
                   
                    
                    for (int i = 0; i < code.Length; i++)
                    {
                        char c = System.Convert.ToChar(code[i]);
                        if (char.IsNumber(c) == true)
                        {
                                
                            if(System.Convert.ToInt32(Code) <= 100)
                            {
                                CodeMsg = "OK";
                            }
                            else if (code.Length > 3)
                            {
                                CodeMsg = "##您的分數超過三位數";
                            }
                            else
                            {
                                CodeMsg = "輸入值大於100";
                            }
                                
                        }
                        else
                        {
                            CodeMsg = "##請輸入整數值";
                            break;
                        }
                    }
                    return CodeMsg;
                                                    
                }
                else
                {
                    CodeMsg = "##您必須輸入";
                }
                return CodeMsg;
            }

        }

       
    }
}
