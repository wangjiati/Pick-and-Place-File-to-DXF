using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using CsvHelper;
using netDxf;
using netDxf.Blocks;
using netDxf.Entities;
using System.Diagnostics;
using Newtonsoft.Json;


using NPOI.SS.UserModel;
using System.Text.RegularExpressions;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;

namespace Pick_and_Place_File_to_DXF
{
    public partial class mainForm : Form
    {
        private Configuration _config;
        public DataTable PickandPlaceTable = new DataTable();
        public DataView  PickandPlaceTableView = new DataView();
        public string PPFunit = "";
        public List<string> Layers = new List<string>();


        public DataTable dxfBlocksTable = new DataTable();
        public DataView  dxfBlocksView = new DataView();
        public List<Block> tableBlocks = new List<Block>();

        public DataTable RulesTable = new DataTable();

        public DxfDocument TBaseMap_Dxf = null;
        public DxfDocument BBaseMap_Dxf = null;


        public DataTable getDXFPickTable = new DataTable();
        public DataView getDXFPickTableView = new DataView();



        //读出的表格格式判断, 判断是否为要求的这些列
        public string[] ColumnName_Designator = { "Designator", "RefDes", "位号"};      //Designator:99se,AD; RefDes:PADS宏
        public string[] ColumnName_Layer = { "Layer", "TB","层" };             //Layer:99SE,AD,PADS宏; TB:99
        public string[] ColumnName_Footprint = { "Footprint", "PartType" ,"封装" };    //Footprint:99se,AD; PartType:PADS宏
        public string[] ColumnName_MidX = { "MidX", "Center-X(mm)", "Center-X(mil)", "X" };   //Center - X(mm):>AD18; MidX:99se,<=AD17; X:PADS宏
        public string[] ColumnName_MidY = { "MidY", "Center-Y(mm)", "Center-Y(mil)", "Y" };   //Center - Y(mm):>AD18; MidY:99se,<=AD17; Y:PADS宏
        public string[] ColumnName_Rotation = { "Rotation", "Orient." ,"角度"};        //Rotation):>99se,AD; Orient:PADS宏
        public string[] ColumnName_Comment = { "Comment", "Value", "Name" ,"型号"};  //Comment:99se,AD; Value:AD; Name:99,PADS
         
        public string[] ColumnName_RefX = { "RefX", "Ref-X(mm)", "Ref-X(mil)" };  //RefX:99se,AD; Ref-X:>AD18;
        public string[] ColumnName_RefY = { "RefY", "Ref-Y(mm)", "Ref-Y(mil)" };  //RefY:99se,AD; Ref-Y:>AD18;
        public string[] ColumnName_PadX = { "PadX", "Pad-X(mm)", "Pad-X(mil)" };  //PadX:99se,AD; Pad-X:>AD18;
        public string[] ColumnName_PadY = { "PadY", "Pad-y(mm)", "Pad-y(mil)" };  //PadY:99se,AD; Pad-Y:>AD18;

        public string[] TLayerName = { "T", "Top", "TopLayer" };
        public string[] BLayerName = { "B", "Bottom", "BottomLayer" };

        public string[] ColumnName_dxfBlockRules = { "Block_Name", "匹配规则(正则表达式)", "匹配规则说明", "规则创建时间", "规则创建人"};    


        //必须的列
        public bool CheckDesignatorOK = false;
        public bool CheckLayerOK = false;
        public bool CheckRotationOK = false;
        public bool CheckMidXOK = false;
        public bool CheckMidYOK = false;

        public bool CheckFootprintOK = false;
        public bool CheckCommentOK = false;
        public bool CheckRefXOK = false;
        public bool CheckRefYOK = false;
        public bool CheckPadXOK = false;
        public bool CheckPadYOK = false;




        public mainForm(string[] args)
        {

            string str1 = System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData);

            if (Directory.Exists(str1 + "\\" + "Pick_and_Place_File_to_DXF") == false)//如果不存,在就创建file文件夹
            {
                Directory.CreateDirectory(str1 + "\\" + "Pick_and_Place_File_to_DXF");
            }


            _config = Configuration.LoadFile(str1 + "\\" + "Pick_and_Place_File_to_DXF" + "\\" + Configuration.CONFIG_FILE);
            _config.FixConfiguration();
            InitializeComponent();


            checkBoxTLayer.Checked =  _config.checkBoxTLayer;
            checkBoxBLayer.Checked = _config.checkBoxBLayer;
            checkBoxBLayerMirrorX.Checked = _config.checkBoxBLayerMirrorX;
            checkBoxBLayerMirrorR.Checked = _config.checkBoxBLayerMirrorR;

            checkBoxDesignatorTXT.Checked = _config.checkBoxDesignatorTXT;
            checkBoxCommentTXT.Checked = _config.checkBoxCommentTXT;
            checkBoxFootprintTXT.Checked = _config.checkBoxFootprintTXT;

            textBoxTXTHeight.Text = _config.textBoxTXTHeight;
            checkBoxTXT180.Checked = _config.checkBoxTXT180;


            if (checkBoxBLayer.Checked)
            {
                checkBoxBLayerMirrorX.Enabled = true;
                checkBoxBLayerMirrorR.Enabled = true;
            }
            else
            {
                checkBoxBLayerMirrorX.Enabled = false;
                checkBoxBLayerMirrorR.Enabled = false;
            }

            if (_config.radioButtonDrawing0603)
            {
                radioButtonNoDrawing.Checked = false;
                radioButtonDrawing0402.Checked = false;
                radioButtonDrawing0603.Checked = true;
            }
            else if (_config.radioButtonDrawing0402)
            {
                radioButtonNoDrawing.Checked = false;
                radioButtonDrawing0402.Checked = true;
                radioButtonDrawing0603.Checked = false;
            }
            else
            {
                radioButtonNoDrawing.Checked = true;
                radioButtonDrawing0402.Checked = false;
                radioButtonDrawing0603.Checked = false;
            }

            if (_config.radioButtonRulesEnabled)
            {
                radioButtonRulesEnabled.Checked = true;
                radioButtonRulesDisabled.Checked = false;
            }
            else
            {
                radioButtonRulesEnabled.Checked = false;
                radioButtonRulesDisabled.Checked = true;
            }
            

            if (_config.radioButtonZeroPoint)
            {
                radioButtonZeroPoint.Checked = true;
                radioButtonNoZeroPoint.Checked = false;
            }
            else
            {
                radioButtonZeroPoint.Checked = false;
                radioButtonNoZeroPoint.Checked = true;
            }

            textBoxDXFBlocksFilePath.Text = _config.textBoxDXFBlocksFilePath;
            textBoxBlocksRulesFilePath.Text = _config.textBoxBlocksRulesFilePath;



            dxfBlocksTable.Columns.Add("Name");
            dxfBlocksTable.Columns.Add("Description");
            dxfBlocksTable.Columns.Add("Layer");
            dxfBlocksTable.Columns.Add("Direction");
            dxfBlocksTable.Columns.Add("Pins");
            dxfBlocksTable.Columns.Add("Polarity");
            dxfBlocksTable.Columns.Add("Path");
            dxfBlocksTable.Columns.Add("匹配规则(正则表达式)"); //注:[规则1][规则2][规则n], 需编辑时修改:MatchingRulesTable.xlsx


            getDXFPickTable.Columns.Add("BlockName");
            getDXFPickTable.Columns.Add("Description");
            getDXFPickTable.Columns.Add("X");
            getDXFPickTable.Columns.Add("Y");
            getDXFPickTable.Columns.Add("Rotation");

            foreach (var ColumnName in ColumnName_dxfBlockRules)
            {
                RulesTable.Columns.Add(ColumnName);
            }


            if (File.Exists(textBoxDXFBlocksFilePath.Text))
            {
                OpendDXF(textBoxDXFBlocksFilePath.Text);
            }
            else
            {
                string binPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                binPath = Path.GetDirectoryName(binPath) + "\\" + textBoxDXFBlocksFilePath.Text;
                if (File.Exists(binPath))
                {
                    OpendDXF(binPath);
                }
            }

            if (File.Exists(textBoxBlocksRulesFilePath.Text))
            {
                openRules(textBoxBlocksRulesFilePath.Text);
            }
            else
            {
                string binPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                binPath = Path.GetDirectoryName(binPath) + "\\" + textBoxBlocksRulesFilePath.Text;
                if (File.Exists(binPath))
                {
                    openRules(binPath);
                }
            }

            //命令行启动?
            if (args.Length == 1)
            {
                if (File.Exists(args[0]))
                {
                    string Extension = Path.GetExtension(args[0]).ToLower();

                    if (Extension == ".csv")
                    {
                        string FilePath = args[0];
                        if (File.Exists(FilePath) && Path.GetExtension(FilePath).ToLower() == ".csv")
                        {
                            OpenPPFcsv(FilePath);
                            OutputDXF();
                        }
                    }
                }
                
            }
        }

        /// <summary>
        /// 打开文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            openFileDialog.Filter = " EXCEL  or csv files (*.csv;*xlsx;*xls)|*.csv;*xlsx;*xls";
            openFileDialog.FilterIndex = 1;
            openFileDialog.Multiselect = false;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                foreach (var FilePath in openFileDialog.FileNames)
                {
                    OpenPPFcsv(FilePath);
                }
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            string FilePath =  textBoxPickAndPlaceFilePath.Text ;
            if (File.Exists(FilePath))
            {
                OpenPPFcsv(FilePath);
            }
        }
        private void OpenPPFcsv(string FilePath)
        {
            try
            {
                DataTable dataTable = new DataTable();
                if (Path.GetExtension(FilePath).ToLower() == ".csv")
                {
                    Stream stream = File.Open(FilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    StreamReader reader = new StreamReader(stream, Encoding.Default);
                    bool table = false;

                    //处理>AD18
                    long Position = 0;
                    string line = reader.ReadLine();
                    Position += System.Text.Encoding.Default.GetByteCount(line) + Encoding.Default.GetByteCount(Environment.NewLine);

                    if (line.Trim().ToLower().Replace(" ", "").Contains("altiumdesignerpickandplacelocations")) //Altium Designer Pick and Place Locations
                    {
                        long startPosition = 0;
                        do
                        {
                            startPosition = Position;
                            line = reader.ReadLine();
                            Position += System.Text.Encoding.Default.GetByteCount(line) + Encoding.Default.GetByteCount(Environment.NewLine); ;

                        } while (line.Trim().ToLower().Replace(" ", "").Contains("designator") == false);

                        reader.DiscardBufferedData();
                        reader.BaseStream.Seek(startPosition, SeekOrigin.Begin);

                        //long Position2 = reader.BaseStream.Position;
                        //string t = reader.ReadToEnd();
                    }
                    else if (line.Trim().ToLower().Replace(" ", "").Contains("\t")) //制表符的CSV格式
                    {
                        table = true;
                        reader.DiscardBufferedData();
                        reader.BaseStream.Seek(0, SeekOrigin.Begin);
                    }
                    else
                    {
                        reader.DiscardBufferedData();
                        reader.BaseStream.Seek(0, SeekOrigin.Begin);
                    }



                    using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                    {
                        csv.Configuration.IgnoreBlankLines = false;
                        if (table)
                        {
                            csv.Configuration.Delimiter = "\t";
                        }
                        // Do any configuration to `CsvReader` before creating CsvDataReader.
                        var dataReader = new CsvDataReader(csv);

                        dataTable.Load(dataReader);
                        //---- 删除空列 空行
                        dataTable = RemoveEmptyRows(dataTable);
                        dataTable = RemvoeEmptyColumns(dataTable);

                        if (dataTable.Rows.Count <= 0)
                        {
                            MessageBox.Show("请确认是否打开了一个空白的工作表");
                            return;
                        }
                    }
                }
                else if(Path.GetExtension(FilePath).ToLower() == ".xls" || Path.GetExtension(FilePath).ToLower() == ".xlsx")
                {
                    dataTable = Excel.ToDataTable(FilePath, "");
                }


                #region 检查读入的文件是否是符合要求的格式
                //----- 初始
                PickandPlaceTable.Clear();
                Layers.Clear();
                PPFunit = "";
                CheckDesignatorOK = false;
                CheckLayerOK = false;
                CheckFootprintOK = false;
                CheckMidXOK = false;
                CheckMidYOK = false;
                CheckRefXOK = false;
                CheckRefYOK = false;
                CheckPadXOK = false;
                CheckPadYOK = false;
                CheckRotationOK = false;
                CheckCommentOK = false;



                //---- 找列
                string temp = "";

                //Designator
                foreach (string ColumnName in ColumnName_Designator)
                {
                    for (int i = 0; i < dataTable.Columns.Count; i++)
                    {
                        string name_c = ColumnName.ToUpper();
                        string name_i = dataTable.Columns[i].ColumnName.Trim().ToUpper().Replace(" ", "");

                        if (String.Compare(name_c, name_i) == 0)
                        {
                            CheckDesignatorOK = true;
                            dataTable.Columns[i].ColumnName = ColumnName_Designator[0];
                            break;
                        }
                    }
                    if (CheckDesignatorOK)
                    {
                        break;
                    }
                }
                if (CheckDesignatorOK == false)
                {
                    temp += "位号列标题未找到. 仅支持这几种:[" + string.Join(",", ColumnName_Designator) + "] \r\n";
                }
                //Layer
                foreach (string ColumnName in ColumnName_Layer)
                {
                    for (int i = 0; i < dataTable.Columns.Count; i++)
                    {
                        if (String.Compare(ColumnName.ToUpper(), dataTable.Columns[i].ColumnName.Trim().ToUpper().Replace(" ", "")) == 0)
                        {
                            CheckLayerOK = true;
                            dataTable.Columns[i].ColumnName = ColumnName_Layer[0];
                            break;
                        }
                    }
                    if (CheckLayerOK)
                    {
                        break;
                    }
                }
                if (CheckLayerOK == false)
                {
                    temp += "层名称列标题未找到. 仅支持这几种:[" + string.Join(",", ColumnName_Layer) + "] \r\n";
                }

                //ColumnName_MidX
                foreach (string ColumnName in ColumnName_MidX)
                {
                    for (int i = 0; i < dataTable.Columns.Count; i++)
                    {
                        if (String.Compare(ColumnName.ToUpper(), dataTable.Columns[i].ColumnName.Trim().ToUpper().Replace(" ", "")) == 0)
                        {
                            CheckMidXOK = true;
                            dataTable.Columns[i].ColumnName = ColumnName_MidX[0];
                            if (string.IsNullOrEmpty(PPFunit))
                            {
                                if ("Center-X(mm)" == ColumnName)
                                {
                                    PPFunit = "mm";
                                }
                                else if ("Center-X(mil)" == ColumnName)
                                {
                                    PPFunit = "mil";
                                }
                            }
                            break;
                        }
                    }
                    if (CheckMidXOK)
                    {
                        break;
                    }
                }
                if (CheckMidXOK == false)
                {
                    temp += "X列标题未找到. 仅支持这几种:[" + string.Join(",", ColumnName_MidX) + "] \r\n";
                }
                //ColumnName_MidY
                foreach (string ColumnName in ColumnName_MidY)
                {
                    for (int i = 0; i < dataTable.Columns.Count; i++)
                    {
                        if (String.Compare(ColumnName.ToUpper(), dataTable.Columns[i].ColumnName.Trim().ToUpper().Replace(" ", "")) == 0)
                        {
                            CheckMidYOK = true;
                            dataTable.Columns[i].ColumnName = ColumnName_MidY[0];
                            break;
                        }
                    }
                    if (CheckMidYOK)
                    {
                        break;
                    }
                }
                if (CheckMidYOK == false)
                {
                    temp += "Y列标题未找到. 仅支持这几种:[" + string.Join(",", ColumnName_MidY) + "] \r\n";
                }
                //ColumnName_Rotation
                foreach (string ColumnName in ColumnName_Rotation)
                {
                    for (int i = 0; i < dataTable.Columns.Count; i++)
                    {
                        if (String.Compare(ColumnName.ToUpper(), dataTable.Columns[i].ColumnName.Trim().ToUpper().Replace(" ", "")) == 0)
                        {
                            CheckRotationOK = true;
                            dataTable.Columns[i].ColumnName = ColumnName_Rotation[0];
                            break;
                        }
                    }
                    if (CheckRotationOK)
                    {
                        break;
                    }
                }
                if (CheckRotationOK == false)
                {
                    temp += "角度列标题未找到. 仅支持这几种:[" + string.Join(",", ColumnName_Rotation) + "] \r\n";
                }


                //以下不是必须存在的列.没有封装名时,使用默认的快名称.
                //Footprint
                foreach (string ColumnName in ColumnName_Footprint)
                {
                    for (int i = 0; i < dataTable.Columns.Count; i++)
                    {
                        if (String.Compare(ColumnName.ToUpper(), dataTable.Columns[i].ColumnName.Trim().ToUpper().Replace(" ", "")) == 0)
                        {
                            CheckFootprintOK = true;
                            dataTable.Columns[i].ColumnName = ColumnName_Footprint[0];
                            break;
                        }
                    }
                    if (CheckFootprintOK)
                    {
                        break;
                    }
                }

                //以下不是必须存在的列,为了方便以后扩展,这里先加上,"Ref X", "Ref Y", "Pad X", "Pad Y",
                //ColumnName_Comment
                foreach (string ColumnName in ColumnName_Comment)
                {
                    for (int i = 0; i < dataTable.Columns.Count; i++)
                    {
                        if (String.Compare(ColumnName.ToUpper(), dataTable.Columns[i].ColumnName.Trim().ToUpper().Replace(" ", "")) == 0)
                        {
                            CheckCommentOK = true;
                            dataTable.Columns[i].ColumnName = ColumnName_Comment[0];
                            break;
                        }
                    }
                    if (CheckCommentOK)
                    {
                        break;
                    }
                }

                //ColumnName_RefX
                foreach (string ColumnName in ColumnName_RefX)
                {
                    for (int i = 0; i < dataTable.Columns.Count; i++)
                    {
                        if (String.Compare(ColumnName.ToUpper(), dataTable.Columns[i].ColumnName.Trim().ToUpper().Replace(" ", "")) == 0)
                        {
                            CheckRefXOK = true;
                            dataTable.Columns[i].ColumnName = ColumnName_RefX[0];
                            break;
                        }
                    }
                    if (CheckRefXOK)
                    {
                        break;
                    }
                }

                //ColumnName_RefY
                foreach (string ColumnName in ColumnName_RefY)
                {
                    for (int i = 0; i < dataTable.Columns.Count; i++)
                    {
                        if (String.Compare(ColumnName.ToUpper(), dataTable.Columns[i].ColumnName.Trim().ToUpper().Replace(" ", "")) == 0)
                        {
                            CheckRefYOK = true;
                            dataTable.Columns[i].ColumnName = ColumnName_RefY[0];
                            break;
                        }
                    }
                    if (CheckRefYOK)
                    {
                        break;
                    }
                }

                //ColumnName_PadX
                foreach (string ColumnName in ColumnName_PadX)
                {
                    for (int i = 0; i < dataTable.Columns.Count; i++)
                    {
                        if (String.Compare(ColumnName.ToUpper(), dataTable.Columns[i].ColumnName.Trim().ToUpper().Replace(" ", "")) == 0)
                        {
                            CheckPadXOK = true;
                            dataTable.Columns[i].ColumnName = ColumnName_PadX[0];
                            break;
                        }
                    }
                    if (CheckPadXOK)
                    {
                        break;
                    }
                }

                //ColumnName_PadY
                foreach (string ColumnName in ColumnName_PadY)
                {
                    for (int i = 0; i < dataTable.Columns.Count; i++)
                    {
                        if (String.Compare(ColumnName.ToUpper(), dataTable.Columns[i].ColumnName.Trim().ToUpper().Replace(" ", "")) == 0)
                        {
                            CheckPadYOK = true;
                            dataTable.Columns[i].ColumnName = ColumnName_PadY[0];
                            break;
                        }
                    }
                    if (CheckPadYOK)
                    {
                        break;
                    }
                }

                //所有必须列都是存在的, 继续操作
                if (string.IsNullOrEmpty(temp))
                {

                    PickandPlaceTable = dataTable.Copy();
                    for (int i = 0; i < PickandPlaceTable.Columns.Count; i++)
                    {
                        PickandPlaceTable.Columns[i].ReadOnly = false;
                    }
                    if (PickandPlaceTable.Columns.Contains("NO.") == false)
                    {
                        PickandPlaceTable.Columns.Add("NO.").SetOrdinal(0);
                    }


                    //PickandPlaceTable.Columns[ColumnName_Layer[0]].ReadOnly = false;
                    //PickandPlaceTable.Columns[ColumnName_MidX[0]].ReadOnly = false;
                    //PickandPlaceTable.Columns[ColumnName_MidY[0]].ReadOnly = false;

                    for (int i = 0; i < PickandPlaceTable.Rows.Count; i++)
                    {
                        //填写序号
                        PickandPlaceTable.Rows[i]["NO."] = (i + 1).ToString();

                        //---- 修正数据 1,层名称
                        bool Aknown = false;
                        string LayerName = PickandPlaceTable.Rows[i][ColumnName_Layer[0]].ToString().Trim();

                        foreach (var TLN in TLayerName)
                        {
                            if (TLN.ToUpper() == LayerName.ToUpper())
                            {
                                PickandPlaceTable.Rows[i][ColumnName_Layer[0]] = TLayerName[0];
                                Aknown = true;
                                break;
                            }
                        }

                        if (Aknown == false)
                        {
                            foreach (var BLN in BLayerName)
                            {
                                if (BLN.ToUpper() == LayerName.ToUpper())
                                {
                                    PickandPlaceTable.Rows[i][ColumnName_Layer[0]] = BLayerName[0];
                                    Aknown = true;
                                    break;
                                }
                            }
                        }
                        /*
                        //不是T,也不是B时, 的处理
                        if (Aknown == false)
                        {

                        }
                        */

                        LayerName = PickandPlaceTable.Rows[i][ColumnName_Layer[0]].ToString();
                        if (Layers.Contains(LayerName) == false)
                        {
                            Layers.Add(LayerName);
                        }


                        //---- 统一单位,转换单位为mm
                        //Mid
                        string MidXStr = PickandPlaceTable.Rows[i][ColumnName_MidX[0]].ToString().Trim().ToLower();
                        string MidYStr = PickandPlaceTable.Rows[i][ColumnName_MidY[0]].ToString().Trim().ToLower();

                        if (string.IsNullOrEmpty(PPFunit))
                        {
                            if (MidXStr.Contains("mm"))
                            {
                                PPFunit = "mm";
                            }
                            else if (MidXStr.Contains("mil"))
                            {
                                PPFunit = "mil";
                            }
                            else //其他情况暂且当他是mm
                            {
                                PPFunit = "mm";
                            }
                        }

                        if (PPFunit == "mil")
                        {
                            char[] charsToTrim = { 'm', 'i', 'l' };
                            MidXStr = MidXStr.TrimEnd(charsToTrim);
                            MidYStr = MidYStr.TrimEnd(charsToTrim);
                        }
                        else if (PPFunit == "mm")
                        {
                            MidXStr = MidXStr.TrimEnd('m');
                            MidYStr = MidYStr.TrimEnd('m');
                        }

                        if (string.IsNullOrEmpty(MidXStr) != true && string.IsNullOrEmpty(MidYStr) != true)
                        {
                            double MidX = Convert.ToDouble(MidXStr);
                            double MidY = Convert.ToDouble(MidYStr);

                            if (PPFunit == "mil")
                            {
                                MidX = MidX * 0.0254;
                                MidY = MidY * 0.0254;
                            }

                            PickandPlaceTable.Rows[i][ColumnName_MidX[0]] = MidX.ToString();
                            PickandPlaceTable.Rows[i][ColumnName_MidY[0]] = MidY.ToString();
                        }
                        else
                        {
                            PickandPlaceTable.Rows[i][ColumnName_MidX[0]] = "";
                            PickandPlaceTable.Rows[i][ColumnName_MidY[0]] = "";
                        }


                        //Ref
                        if (CheckRefXOK && CheckRefYOK)
                        {
                            PickandPlaceTable.Columns[ColumnName_RefX[0]].ReadOnly = false;
                            PickandPlaceTable.Columns[ColumnName_RefY[0]].ReadOnly = false;

                            string RefXStr = PickandPlaceTable.Rows[i][ColumnName_RefX[0]].ToString().Trim().ToLower();
                            string RefYStr = PickandPlaceTable.Rows[i][ColumnName_RefY[0]].ToString().Trim().ToLower();

                            if (PPFunit == "mil")
                            {
                                char[] charsToTrim = { 'm', 'i', 'l' };
                                RefXStr = RefXStr.TrimEnd(charsToTrim);
                                RefYStr = RefYStr.TrimEnd(charsToTrim);
                            }
                            else if (PPFunit == "mm")
                            {
                                RefXStr = RefXStr.TrimEnd('m');
                                RefYStr = RefYStr.TrimEnd('m');
                            }


                            if (string.IsNullOrEmpty(RefXStr) != true && string.IsNullOrEmpty(RefYStr) != true)
                            {
                                double RefX = Convert.ToDouble(RefXStr);
                                double RefY = Convert.ToDouble(RefYStr);
                                if (PPFunit == "mil")
                                {
                                    RefX = RefX * 0.0254;
                                    RefY = RefY * 0.0254;
                                }
                                PickandPlaceTable.Rows[i][ColumnName_RefX[0]] = RefX.ToString();
                                PickandPlaceTable.Rows[i][ColumnName_RefY[0]] = RefY.ToString();
                            }
                            else
                            {
                                PickandPlaceTable.Rows[i][ColumnName_RefX[0]] = "";
                                PickandPlaceTable.Rows[i][ColumnName_RefY[0]] = "";
                            }
                        }
                        //Pad
                        if (CheckPadXOK && CheckPadYOK)
                        {
                            PickandPlaceTable.Columns[ColumnName_PadX[0]].ReadOnly = false;
                            PickandPlaceTable.Columns[ColumnName_PadY[0]].ReadOnly = false;

                            string PadXStr = PickandPlaceTable.Rows[i][ColumnName_PadX[0]].ToString().Trim().ToLower();
                            string PadYStr = PickandPlaceTable.Rows[i][ColumnName_PadY[0]].ToString().Trim().ToLower();

                            if (PPFunit == "mil")
                            {
                                char[] charsToTrim = { 'm', 'i', 'l' };
                                PadXStr = PadXStr.TrimEnd(charsToTrim);
                                PadYStr = PadYStr.TrimEnd(charsToTrim);
                            }
                            else if (PPFunit == "mm")
                            {
                                PadXStr = PadXStr.TrimEnd('m');
                                PadYStr = PadYStr.TrimEnd('m');
                            }


                            if (string.IsNullOrEmpty(PadXStr) != true && string.IsNullOrEmpty(PadYStr) != true)
                            {
                                double PadX = Convert.ToDouble(PadXStr);
                                double PadY = Convert.ToDouble(PadYStr);
                                if (PPFunit == "mil")
                                {
                                    PadX = PadX * 0.0254;
                                    PadY = PadY * 0.0254;
                                }
                                PickandPlaceTable.Rows[i][ColumnName_PadX[0]] = PadX.ToString();
                                PickandPlaceTable.Rows[i][ColumnName_PadY[0]] = PadY.ToString();
                            }
                            else
                            {
                                PickandPlaceTable.Rows[i][ColumnName_PadX[0]] = "";
                                PickandPlaceTable.Rows[i][ColumnName_PadY[0]] = "";
                            }
                        }

                        //---- 清除无效行数据
                        //暂时先不写清除无效行
                    }

                    PickandPlaceTableView = PickandPlaceTable.DefaultView;
                    dataGridViewPickandPlaceTableView.DataSource = PickandPlaceTableView;
                    //dataGridViewPickandPlaceTableView.RowHeadersWidth = 60;
                    textBoxPickAndPlaceFilePath.Text = FilePath;

                    //报错
                    if (Layers.Count > 2 && Layers.Count < 10)
                    {
                        MessageBox.Show("层名称应该是不对,目前仅支持T,B两层,找到了这么多种层名称:" + string.Join(";", Layers));
                    }
                    else if (Layers.Count > 2)
                    {
                        MessageBox.Show("层名称错误,目前仅支持T,B两层,找到了层名称数量:" + Layers.Count.ToString());
                    }


                    //----界面配置
                    //T层存在时
                    if (Layers.Contains(TLayerName[0]))
                    {
                        checkBoxTLayer.Enabled = true;
                    }
                    else
                    {
                        checkBoxTLayer.Enabled = false;
                    }
                    //B层存在时
                    if (Layers.Contains(BLayerName[0]))
                    {
                        checkBoxBLayer.Enabled = true;

                        checkBoxBLayerMirrorX.Enabled = true;
                        checkBoxBLayerMirrorR.Enabled = true;
                    }
                    else
                    {
                        checkBoxBLayer.Enabled = false;

                        checkBoxBLayerMirrorX.Enabled = false;
                        checkBoxBLayerMirrorR.Enabled = false;
                    }

                    if (CheckFootprintOK)
                    {
                        checkBoxFootprintTXT.Enabled = true;
                    }
                    else
                    {
                        checkBoxFootprintTXT.Enabled = false;
                    }

                    if (CheckCommentOK)
                    {
                        checkBoxCommentTXT.Enabled = true;
                    }
                    else
                    {
                        checkBoxCommentTXT.Enabled = false;
                    }

                    checkBoxUsinFixRotation.Enabled = false;
                    checkBoxUsinFixRotation.Checked = false;
                }
                else
                {
                    MessageBox.Show("仅支持固定格式的坐标文件 \r\n\r\n" + temp);
                }

                #endregion

            }
            catch (Exception e1)
            {
                MessageBox.Show(e1.Message + "\r\n 由于CSV没有严格的编码格式定义,暂时做不到支持所有情况,  请尝试使用Excel打开,然后另存为CSV试试 ");
                //throw;
            }
        }



        //DataTable 删除空行
        protected DataTable RemoveEmptyRows(DataTable dt)
        {
            List<DataRow> removelist = new List<DataRow>();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                bool rowdataisnull = true;
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    if (!string.IsNullOrEmpty(dt.Rows[i][j].ToString().Trim()))
                    {

                        rowdataisnull = false;
                        break;  //只要有一个单元格网是用内容，则不用再删行
                    }
                }
                if (rowdataisnull)
                {
                    removelist.Add(dt.Rows[i]);
                }

            }
            for (int i = 0; i < removelist.Count; i++)
            {
                dt.Rows.Remove(removelist[i]);
            }
            return dt;
        }

        //DataTable 删除空列, 注意只要列中没有内容就会删除列，如果是列中确实是空内容，又要保留，那就不能调用此方法了。
        protected DataTable RemvoeEmptyColumns(DataTable dt)
        {
            List<DataColumn> removelist = new List<DataColumn>();
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                bool rowdataisnull = true;
                for (int j = 0; j < dt.Rows.Count; j++)
                {
                    if (!string.IsNullOrEmpty(dt.Rows[j][i].ToString().Trim()))
                    {
                        rowdataisnull = false;
                        break;
                    }
                }
                if (rowdataisnull)
                {
                    removelist.Add(dt.Columns[i]);
                }
            }
            for (int i = 0; i < removelist.Count; i++)
            {
                dt.Columns.Remove(removelist[i]);
            }
            return dt;
        }

        //DataTable 删掉为空的列，且列名中包含指定关键字称
        protected DataTable RemvoeEmptyColumnsAndColumnName(DataTable dt, string str)
        {
            List<DataColumn> removelist = new List<DataColumn>();
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                bool rowdataisnull = true;
                for (int j = 0; j < dt.Rows.Count; j++)
                {
                    if (!string.IsNullOrEmpty(dt.Rows[j][i].ToString().Trim()))
                    {
                        rowdataisnull = false;
                        break;
                    }
                }
                if (rowdataisnull)
                {
                    if (dt.Columns[i].ColumnName.Contains(str)) //不为空且包含关键字，大小写敏感的
                    {
                        removelist.Add(dt.Columns[i]);
                    }
                }
            }
            for (int i = 0; i < removelist.Count; i++)
            {
                dt.Columns.Remove(removelist[i]);
            }
            return dt;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            openFileDialog.Filter = "DXF files (*.dxf)|*.dxf|All files (*.*)|*.*";
            openFileDialog.FilterIndex = 1;
            openFileDialog.Multiselect = true;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                foreach (var FilePath in openFileDialog.FileNames)
                {
                    if (Path.GetExtension(FilePath).ToLower() == ".dxf")
                    {
                        OpendDXF(FilePath);

                        //重新打开规则
                        if (dxfBlocksTable.Rows.Count > 0 && string.IsNullOrWhiteSpace(textBoxBlocksRulesFilePath.Text) == false)
                        {
                            openRules(textBoxBlocksRulesFilePath.Text);
                        }
                    }

                }
            }
        }

        void OpendDXF(string FilePath)
        {
            if (File.Exists(FilePath))
            {
                tableBlocks.Clear();
                dxfBlocksTable.Clear();

                DxfDocument doc = DxfDocument.Load(FilePath);

                foreach (Block block in doc.Blocks)
                {
                    if (block.IsForInternalUseOnly == false)
                    {
                        if (string.IsNullOrEmpty(block.Name) == false) //块名称不能为空
                        {
                            string Name = "";
                            string Description = "";
                            string Layer = "";
                            string Direction = "";
                            string Pins = "";
                            string Polarity = "";
                            string Path = FilePath;

                            foreach (var Entitie in block.Entities)
                            {
                                if (Entitie.Type == EntityType.Circle)
                                {
                                    Circle pin1 = (Circle)Entitie;
                                    if (string.Compare(pin1.Layer.Name, "Pin1") == 0 && (0.03 >= pin1.Radius && pin1.Radius >= 0.015))
                                    {

                                        double angleOfLine = Math.Atan2((pin1.Center.Y), (pin1.Center.X)) * 180 / Math.PI;  //(y,x)

                                        angleOfLine = angleOfLine < 0 ? angleOfLine + 360 : angleOfLine;

                                        Direction = angleOfLine.ToString();

                                        //if (337.5 >= angleOfLine && 22.5 > angleOfLine)
                                        //{
                                        //    Direction = "0";
                                        //}
                                        //else if (67.5 >= angleOfLine && angleOfLine > 22.5)
                                        //{
                                        //    Direction = "1";
                                        //}
                                        //else if (112.5 >= angleOfLine && angleOfLine > 67.5)
                                        //{
                                        //    Direction = "2";
                                        //}
                                        //else if (157.5 >= angleOfLine && angleOfLine > 112.5)
                                        //{
                                        //    Direction = "3";
                                        //}
                                        //else if (202.5 >= angleOfLine && angleOfLine > 157.5)
                                        //{
                                        //    Direction = "4";
                                        //}
                                        //else if (247.5 >= angleOfLine && angleOfLine > 202.5)
                                        //{
                                        //    Direction = "5";
                                        //}
                                        //else if (292.5 >= angleOfLine && angleOfLine > 247.5)
                                        //{
                                        //    Direction = "6";
                                        //}
                                        //else //if (337.5 >= angleOfLine && angleOfLine > 292.5)
                                        //{
                                        //    Direction = "7";
                                        //}
                                    }
                                }
                            }

                            Name = block.Name;
                            Description = block.Description;
                            Layer = block.Layer.Name;

                            dxfBlocksTable.Rows.Add(Name, Description, Layer, Direction, Pins, Polarity, Path);
                            tableBlocks.Add(block);
                        }

                        
                    }
                }

                dxfBlocksView = dxfBlocksTable.DefaultView;
                dataGridViewDXFBlocksTable.DataSource = dxfBlocksView;
                textBoxDXFBlocksFilePath.Text = FilePath;


                string binPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                binPath = Path.GetDirectoryName(binPath);


                string BlocksFilePath; 

                if (binPath == Path.GetDirectoryName(FilePath))
                {
                    BlocksFilePath = Path.GetFileName(FilePath);
                }
                else
                {
                    BlocksFilePath = FilePath;
                }


                if (_config.textBoxDXFBlocksFilePath != BlocksFilePath)
                {
                    _config.textBoxDXFBlocksFilePath = BlocksFilePath;
                    Configuration.Save(_config);
                }

                

            }
        }

        //生成DXF图
        private void button3_Click(object sender, EventArgs e)
        {
            OutputDXF();
        }

        public void OutputDXF()
        {

            string okDXF = "";

            if (PickandPlaceTable.Rows.Count > 0)
            {
                double TXTHeight = 0.8;
                Stopwatch watch = new Stopwatch();
                watch.Start();


                if (string.IsNullOrEmpty(textBoxTXTHeight.Text.Trim().Replace(" ", "")) == false)
                {
                    try
                    {
                        TXTHeight = Convert.ToDouble(textBoxTXTHeight.Text.Trim().Replace(" ", ""));
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("文字高度输入错误,仅支持数字");
                        //throw;
                    }

                }
                if (TXTHeight > 100 && TXTHeight < 0.1)
                {
                    if (MessageBox.Show("文字高度可能合理,现在仅支持 0.1-100", "错误提示,是否使用默认值 0.8mm继续?", MessageBoxButtons.OKCancel) == DialogResult.Cancel)
                    {
                        return;
                    }
                    else
                    {
                        TXTHeight = 0.8;
                    }

                }

                if (PickandPlaceTable.Columns.Contains("生成结果") == false)
                {
                    PickandPlaceTable.Columns.Add("生成结果");
                }
                if (PickandPlaceTable.Columns.Contains("库名称匹配规则") == false)
                {
                    PickandPlaceTable.Columns.Add("库名称匹配规则");
                }

                foreach (string Lay in Layers)
                {
                    if (Lay == TLayerName[0] && checkBoxTLayer.Enabled == false && checkBoxTLayer.Checked == false)
                    {
                        continue;
                    }
                    else if (Lay == BLayerName[0] && checkBoxBLayer.Enabled == false && checkBoxBLayer.Checked == false)
                    {
                        continue;
                    }

                    // create a document
                    DxfDocument doc = new DxfDocument();
                    if (Lay == BLayerName[0] && BBaseMap_Dxf != null)
                    {
                        doc = BBaseMap_Dxf;
                    }
                    else if (Lay == TLayerName[0] && TBaseMap_Dxf != null)
                    {
                        doc = TBaseMap_Dxf;
                    }
                    
                    if (radioButtonZeroPoint.Checked)
                    {
                        Block ZeroP = new Block("00");
                        ZeroP = (Block)ZeroPoint().Clone();
                        ZeroP.Name = "ZeroPoint";

                        Insert Insert = new Insert(ZeroP, new Vector3(0, 0, 0));
                        Insert.Rotation = 0;
                        doc.AddEntity(Insert);
                    }
                    for (int i = 0; i < PickandPlaceTable.Rows.Count; i++)
                    {
                        string Layer = PickandPlaceTable.Rows[i][ColumnName_Layer[0]].ToString().Trim();
                        if (Layer == Lay)
                        {
                            double X, Y, R;
                            string NO = PickandPlaceTable.Rows[i]["NO."].ToString();
                            string Designator = PickandPlaceTable.Rows[i][ColumnName_Designator[0]].ToString().Trim();
                            string MidX = PickandPlaceTable.Rows[i][ColumnName_MidX[0]].ToString().Trim();
                            string MidY = PickandPlaceTable.Rows[i][ColumnName_MidY[0]].ToString().Trim();
                            string Rotation;

                            if (PickandPlaceTable.Columns.Contains("FixRotation") == true && checkBoxUsinFixRotation.Checked)
                            {
                                Rotation = PickandPlaceTable.Rows[i]["FixRotation"].ToString().Trim();
                            }
                            else
                            {
                                Rotation = PickandPlaceTable.Rows[i][ColumnName_Rotation[0]].ToString().Trim();
                            }

                            string Comment = "";
                            string Footprint = "";

                            if (CheckCommentOK)
                            {
                                Comment = PickandPlaceTable.Rows[i][ColumnName_Comment[0]].ToString().Trim();
                            }
                            if (CheckFootprintOK)
                            {
                                Footprint = PickandPlaceTable.Rows[i][ColumnName_Footprint[0]].ToString().Trim();
                            }



                            if (string.IsNullOrEmpty(MidX) == false && string.IsNullOrEmpty(MidY) == false &&
                                string.IsNullOrEmpty(Rotation) == false && string.IsNullOrEmpty(Layer) == false)
                            {
                                try
                                {
                                    X = Convert.ToDouble(MidX);
                                    Y = Convert.ToDouble(MidY);
                                    R = Convert.ToDouble(Rotation);

                                    if (checkBoxBLayerMirrorX.Checked && Lay == BLayerName[0]) //B层是否镜像
                                    {
                                        X = (X - X * 2);
                                    }
                                    if (checkBoxBLayerMirrorR.Checked && Lay == BLayerName[0])
                                    {
                                        R = R % 360;
                                        if (R <= 180)
                                        {
                                            R = 180 - R;
                                        }
                                        else
                                        {
                                            R = 540 - R;
                                        }
                                    }

                                    // ---- 画图
                                    bool isDrawing = false;
                                    Block block = new Block("00");

                                    if (string.IsNullOrEmpty(Footprint) == true && radioButtonDrawing0402.Checked)
                                    {
                                        block = (Block)R0402().Clone();
                                        isDrawing = true;
                                    }
                                    else if (string.IsNullOrEmpty(Footprint) == true && radioButtonDrawing0603.Checked)
                                    {
                                        block = (Block)R0603().Clone();
                                        isDrawing = true;
                                    }
                                    else if (string.IsNullOrEmpty(Footprint) == false)
                                    {
                                        foreach (Block Blo in tableBlocks)
                                        {
                                            if (Footprint == Blo.Name)
                                            {
                                                block = (Block)Blo.Clone();// new Block(Designator + "#" + Footprint + "#" + i.ToString());

                                                PickandPlaceTable.Rows[i]["库名称匹配规则"] = "名称相同";
                                                PickandPlaceTable.Rows[i]["生成结果"] = "OK";
                                                isDrawing = true;
                                                break;
                                            }
                                            else if (radioButtonRulesEnabled.Checked && RulesTable.Rows.Count >0)
                                            {
                                                DataView RulesView = RulesTable.DefaultView;
                                                RulesView.RowFilter = "Block_Name = '" + Blo.Name + "'";
                                                DataTable RowFiltertable = RulesView.ToTable();

                                                for (int r = 0; r < RowFiltertable.Rows.Count; r++)
                                                {
                                                    string Rule = RowFiltertable.Rows[r]["匹配规则(正则表达式)"].ToString();
                                                    if (Regex.IsMatch(Footprint, Rule))
                                                    {
                                                        block = (Block)Blo.Clone();// new Block(Designator + "#" + Footprint + "#" + i.ToString());

                                                        isDrawing = true;

                                                        PickandPlaceTable.Rows[i]["库名称匹配规则"] = "正则匹配" + Rule;
                                                        PickandPlaceTable.Rows[i]["生成结果"] = "OK";
                                                        break;
                                                    }
                                                }
                                                if (isDrawing)
                                                {
                                                    break;
                                                }
                                            }
                                        }
                                        if (isDrawing == false)
                                        {
                                            if (radioButtonDrawing0402.Checked)
                                            {
                                                block = (Block)R0402().Clone();
                                                PickandPlaceTable.Rows[i]["库名称匹配规则"] = "找不到,强制使用0402";
                                                PickandPlaceTable.Rows[i]["生成结果"] = "OK";
                                                isDrawing = true;
                                            }
                                            else if (radioButtonDrawing0603.Checked)
                                            {
                                                block = (Block)R0603().Clone();
                                                PickandPlaceTable.Rows[i]["库名称匹配规则"] = "找不到,强制使用0603";
                                                PickandPlaceTable.Rows[i]["生成结果"] = "OK";
                                                isDrawing = true;
                                            }
                                            else
                                            {
                                                PickandPlaceTable.Rows[i]["库名称匹配规则"] = "找不到,不画";
                                                PickandPlaceTable.Rows[i]["生成结果"] = "NG";
                                            }
                                        }
                                    }
                                    if (isDrawing)
                                    {
                                        string name = "";
                                        string Description = "";

                                        name = Designator + "#" + NO + "#" + Layer;

                                        char[] characters = { '\\', '<', '>', '/', '?', '"', ':', ';', '*', '|', ',', '=', '`' };  //\<>/?":;*|,=`
                                        name = StripCharacters(name, characters, '_');

                                        Description = "[" + Comment + "]_[" + Footprint + "]";


                                        block.Name = name;
                                        block.Description = Description;

                                        Insert nestedInsert = new Insert(block, new Vector3(X, Y, 0));
                                        nestedInsert.Rotation = R;
                                        doc.AddEntity(nestedInsert);
                                    }


                                    // ---- 写字
                                    if (checkBoxDesignatorTXT.Checked && checkBoxDesignatorTXT.Enabled)
                                    {
                                        double TXTlength = TXTHeight * 0.85 * Designator.Length;
                                        double TXT_X = X - TXTlength / 2;
                                        double TXT_Y = Y - TXTHeight / 2;
                                        double TXT_R = checkBoxTXT180.Checked ? R % 180 : R;

                                        Text DesignatorText = new Text(Designator, spinPoint(X, Y, TXT_R, TXT_X, TXT_Y), TXTHeight);
                                        //DesignatorText.Lineweight = Lineweight.W20; // 0.2 mm
                                        DesignatorText.Rotation = TXT_R;
                                        DesignatorText.Layer = new netDxf.Tables.Layer("DesignatorText");
                                        doc.AddEntity(DesignatorText);
                                    }

                                    if (checkBoxCommentTXT.Checked && checkBoxCommentTXT.Enabled && string.IsNullOrEmpty(Comment) == false)
                                    {
                                        double TXTlength = TXTHeight * 0.85 * Comment.Length;
                                        double TXT_X = X - TXTlength / 2;
                                        double TXT_Y = Y - TXTHeight / 2;
                                        double TXT_R = checkBoxTXT180.Checked ? R % 180 : R;

                                        Text CommentText = new Text(Comment, spinPoint(X, Y, TXT_R, TXT_X, TXT_Y), TXTHeight);
                                        //CommentText.Lineweight = Lineweight.W20; // 0.2 mm
                                        CommentText.Layer = new netDxf.Tables.Layer("CommentText");
                                        CommentText.Rotation = TXT_R;
                                        doc.AddEntity(CommentText);
                                    }

                                    if (checkBoxFootprintTXT.Checked && checkBoxFootprintTXT.Enabled && string.IsNullOrEmpty(Footprint) == false)
                                    {
                                        double TXTlength = TXTHeight * 0.85 * Footprint.Length;
                                        double TXT_X = X - TXTlength / 2;
                                        double TXT_Y = Y - TXTHeight / 2;
                                        double TXT_R = checkBoxTXT180.Checked ? R % 180 : R;

                                        Text FootprintText = new Text(Footprint, spinPoint(X, Y, TXT_R, TXT_X, TXT_Y), TXTHeight);
                                        //FootprintText.Lineweight = Lineweight.W20; // 0.2 mm
                                        FootprintText.Layer = new netDxf.Tables.Layer("FootprintText");
                                        FootprintText.Rotation = TXT_R;
                                        doc.AddEntity(FootprintText);
                                    }
                                }
                                catch (Exception)
                                {
                                    throw;
                                }
                            }
                        }
                    }
                    string path = Path.GetDirectoryName(textBoxPickAndPlaceFilePath.Text) + "\\" +
                                    Path.GetFileNameWithoutExtension(textBoxPickAndPlaceFilePath.Text) + "_" +
                                    Lay +
                                    ".dxf";

                    doc.Save(path);
                    okDXF += path + "\r\n";
                }

                watch.Stop();

                if (string.IsNullOrEmpty(okDXF) == false)
                {
                    MessageBox.Show("转DXF已完成,用时:" + (watch.ElapsedMilliseconds / 1000.0).ToString() + "秒,生成文件存放位置:\r\n" + okDXF);
                }
                else
                {
                    MessageBox.Show("发生错误,未有DXF文件生成:");
                }
            }
            else
            {
                MessageBox.Show("没有需要制图数据");
            }

        }
        private void button6_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            openFileDialog.Filter = "DXF files (*.dxf)|*.dxf|All files (*.*)|*.*";
            openFileDialog.FilterIndex = 1;
            openFileDialog.Multiselect = true;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                foreach (var FilePath in openFileDialog.FileNames)
                {
                    if (Path.GetExtension(FilePath).ToLower() == ".dxf")
                    {
                        getDXFPick(FilePath);
                    }
                }
            }
        }
        private void getDXFPick(string FilePath)
        {
            if (File.Exists(FilePath))
            {
                try
                {
                    DxfDocument doc = DxfDocument.Load(FilePath);


                    if (doc.Inserts.Count() > 0)
                    {
                        getDXFPickTable.Dispose();
                        getDXFPickTable = new DataTable();

                        getDXFPickTable.Columns.Add("BlockName");
                        getDXFPickTable.Columns.Add("Description");
                        getDXFPickTable.Columns.Add("X");
                        getDXFPickTable.Columns.Add("Y");
                        getDXFPickTable.Columns.Add("Rotation");

                        foreach (Insert Insert in doc.Inserts)
                        {
                            double X, Y, R;
                            R = Insert.Rotation;
                            X = Insert.Position.X;
                            Y = Insert.Position.Y;

                            getDXFPickTable.Rows.Add(Insert.Block.Name, Insert.Block.Description, X.ToString(), Y.ToString(), R.ToString());
                        }

                        getDXFPickTableView = getDXFPickTable.DefaultView;
                        dataGridViewgetDXFPick.DataSource = getDXFPickTableView;
                        dataGridViewgetDXFPick.RowHeadersWidth = 60;
                        textBoxgetDXFPick.Text = FilePath;
                        button8.Enabled = true;
                        button7.Enabled = true;

                    }
                    else
                    {
                        MessageBox.Show("打开的DXF文件中没有 Block 信息");
                        return;
                    }
                    
                }
                catch (Exception)
                {

                    throw;
                }
                

            }
        }


        /// <summary>
        /// 返回坐标点沿中心店逆时针旋转后的坐标点
        /// xx= (x - dx)*cos(a) - (y - dy)*sin(a) + dx 
        /// yy= (x - dx)*sin(a) + (y - dy)*cos(a) + dy ;
        /// </summary>
        /// <param name="centerX"></param>
        /// <param name="centerY"></param>
        /// <param name="angle"></param>
        /// <param name="X"></param>
        /// <param name="Y"></param>
        /// <returns></returns>
        public Vector2 spinPoint(double centerX, double centerY, double angle,  double X,  double Y)
        {
            Vector2 vector2 = new Vector2();

            double angleHude = angle * Math.PI / 180;/*角度变成弧度*/
            vector2.X = (X - centerX) * Math.Cos(angleHude) - (Y - centerY) * Math.Sin(angleHude) + centerX;
            vector2.Y = (X - centerX) * Math.Sin(angleHude) + (Y - centerY) * Math.Cos(angleHude) + centerY;
        
            return vector2;

        }

        public Block R0603 ()
        {
            Block block = new Block("0603");
            Line lineL = new Line(new Vector2(-0.8, -0.4), new Vector2(-0.8, 0.4));
            lineL.Layer = new netDxf.Tables.Layer("ComponentBody");
            Line lineT = new Line(new Vector2(-0.8, 0.4), new Vector2(0.8, 0.4));
            lineT.Layer = new netDxf.Tables.Layer("ComponentBody");
            Line lineR = new Line(new Vector2(0.8, 0.4), new Vector2(0.8, -0.4));
            lineR.Layer = new netDxf.Tables.Layer("ComponentBody");
            Line lineB = new Line(new Vector2(-0.8, -0.4), new Vector2(0.8, -0.4));
            lineB.Layer = new netDxf.Tables.Layer("ComponentBody");
            Circle circleP1 = new Circle(new Vector2(-0.6, 0),0.1);
            circleP1.Layer = new netDxf.Tables.Layer("ComponentBody");
            Circle circleZ = new Circle(new Vector2(0, 0), 0.05);
            circleZ.Layer = new netDxf.Tables.Layer("ComponentBody");


            block.Entities.Add(lineL);
            block.Entities.Add(lineT);
            block.Entities.Add(lineR);
            block.Entities.Add(lineB);
            block.Entities.Add(circleP1);
            block.Entities.Add(circleZ);
            return block;
        }

        public Block R0402()
        {
            Block block = new Block("0402");
            Line lineL = new Line(new Vector2(-0.5, -0.25), new Vector2(-0.5,  0.25));
            lineL.Layer = new netDxf.Tables.Layer("ComponentBody");
            Line lineT = new Line(new Vector2(-0.5,  0.25), new Vector2( 0.5,  0.25));
            lineT.Layer = new netDxf.Tables.Layer("ComponentBody");
            Line lineR = new Line(new Vector2( 0.5,  0.25), new Vector2( 0.5, -0.25));
            lineR.Layer = new netDxf.Tables.Layer("ComponentBody");
            Line lineB = new Line(new Vector2(-0.5, -0.25), new Vector2( 0.5, -0.25));
            lineB.Layer = new netDxf.Tables.Layer("ComponentBody");
            Circle circleP1 = new Circle(new Vector2(-0.38, 0), 0.1);
            circleP1.Layer = new netDxf.Tables.Layer("ComponentBody");
            Circle circleZ = new Circle(new Vector2(0, 0), 0.05);
            circleZ.Layer = new netDxf.Tables.Layer("ComponentBody");
            block.Entities.Add(lineL);
            block.Entities.Add(lineT);
            block.Entities.Add(lineR);
            block.Entities.Add(lineB);
            block.Entities.Add(circleP1);
            block.Entities.Add(circleZ);
            return block;
        }

        public Block ZeroPoint()
        {
            Block block = new Block("ZeroPoint");
            Circle circleZ = new Circle(new Vector2(0, 0), 0.5);
            circleZ.Layer = new netDxf.Tables.Layer("ZeroPoint");
            block.Entities.Add(circleZ);
            return block;
        }

        private void checkBoxBLayer_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxBLayer.Checked)
            {
                checkBoxBLayerMirrorX.Enabled = true;
                checkBoxBLayerMirrorR.Enabled = true;
            }
            else
            {
                checkBoxBLayerMirrorX.Enabled = false;
                checkBoxBLayerMirrorR.Enabled = false;
            }
            if (_config.checkBoxBLayer != checkBoxBLayer.Checked)
            {
                _config.checkBoxBLayer = checkBoxBLayer.Checked;
                Configuration.Save(_config);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            AboutBox1 about = new AboutBox1();
            about.Show();
        }

        //显示行号
        private void dataGridViewPickandPlaceTableView_RowStateChanged(object sender, DataGridViewRowStateChangedEventArgs e)
        {
            e.Row.HeaderCell.Value = string.Format("{0}", e.Row.Index + 1);
        }
        public static unsafe string StripCharacters(string s , char [] chr , char replace)
        {
            int len = s.Length;
            char* newChars = stackalloc char[len];
            char* currentChar = newChars;

            for (int i = 0; i < len; ++i)
            {
                char c = s[i];

                if (chr.Contains(c) == false )
                {
                    *currentChar++ = c;
                }
                else
                {
                    *currentChar++ = replace;
                }
            }
            return new string(newChars, 0, (int)(currentChar - newChars));
        }

        private void checkBoxTLayer_CheckedChanged(object sender, EventArgs e)
        {
            if (_config.checkBoxTLayer != checkBoxTLayer.Checked)
            {
                _config.checkBoxTLayer = checkBoxTLayer.Checked;
                Configuration.Save(_config);
            }
        }

        private void checkBoxBLayerMirrorX_CheckedChanged(object sender, EventArgs e)
        {
            if (_config.checkBoxBLayerMirrorX != checkBoxBLayerMirrorX.Checked)
            {
                _config.checkBoxBLayerMirrorX = checkBoxBLayerMirrorX.Checked;
                Configuration.Save(_config);
            }
        }

        private void checkBoxDesignatorTXT_CheckedChanged(object sender, EventArgs e)
        {
            if (_config.checkBoxDesignatorTXT != checkBoxDesignatorTXT.Checked)
            {
                _config.checkBoxDesignatorTXT = checkBoxDesignatorTXT.Checked;
                Configuration.Save(_config);
            }
        }

        private void checkBoxCommentTXT_CheckedChanged(object sender, EventArgs e)
        {
            if (_config.checkBoxCommentTXT != checkBoxCommentTXT.Checked)
            {
                _config.checkBoxCommentTXT = checkBoxCommentTXT.Checked;
                Configuration.Save(_config);
            } 
        }

        private void checkBoxFootprintTXT_CheckedChanged(object sender, EventArgs e)
        {
            if (_config.checkBoxFootprintTXT != checkBoxFootprintTXT.Checked)
            {
                _config.checkBoxFootprintTXT = checkBoxFootprintTXT.Checked;
                Configuration.Save(_config);
            }
        }

        private void checkBoxTXT180_CheckedChanged(object sender, EventArgs e)
        {
            if (_config.checkBoxTXT180 != checkBoxTXT180.Checked)
            {
                _config.checkBoxTXT180 = checkBoxTXT180.Checked;
                Configuration.Save(_config);
            }
        }

        private void radioButtonDrawing0603_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButtonDrawing0603.Checked)
            {
                if (_config.radioButtonNoDrawing != radioButtonNoDrawing.Checked || 
                    _config.radioButtonDrawing0402 != radioButtonDrawing0402.Checked ||
                    _config.radioButtonDrawing0603 != radioButtonDrawing0603.Checked)
                {
                    _config.radioButtonNoDrawing = false;
                    _config.radioButtonDrawing0402 = false;
                    _config.radioButtonDrawing0603 = true;
                    Configuration.Save(_config);
                }
                
            }
        }

        private void radioButtonDrawing0402_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButtonDrawing0402.Checked)
            {
                if (_config.radioButtonNoDrawing != radioButtonNoDrawing.Checked ||
                    _config.radioButtonDrawing0402 != radioButtonDrawing0402.Checked ||
                    _config.radioButtonDrawing0603 != radioButtonDrawing0603.Checked)
                {
                    _config.radioButtonNoDrawing = false;
                    _config.radioButtonDrawing0402 = true;
                    _config.radioButtonDrawing0603 = false;
                    Configuration.Save(_config);
                }
                    
            }
        }

        private void radioButtonNoDrawing_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButtonNoDrawing.Checked)
            {
                if (_config.radioButtonNoDrawing != radioButtonNoDrawing.Checked ||
                    _config.radioButtonDrawing0402 != radioButtonDrawing0402.Checked ||
                    _config.radioButtonDrawing0603 != radioButtonDrawing0603.Checked)
                {
                    _config.radioButtonNoDrawing = true;
                    _config.radioButtonDrawing0402 = false;
                    _config.radioButtonDrawing0603 = false;
                    Configuration.Save(_config);
                }
                    
            }
        }

        private void radioButtonZeroPoint_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButtonZeroPoint.Checked)
            {
                if (_config.radioButtonZeroPoint != radioButtonZeroPoint.Checked ||
                    _config.radioButtonNoZeroPoint != radioButtonNoZeroPoint.Checked )
                {
                    _config.radioButtonZeroPoint = true;
                    _config.radioButtonNoZeroPoint = false;
                    Configuration.Save(_config);
                }
                   
            }
        }

        private void radioButtonNoZeroPoint_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButtonZeroPoint.Checked)
            {
                if (_config.radioButtonZeroPoint != radioButtonZeroPoint.Checked ||
                    _config.radioButtonNoZeroPoint != radioButtonNoZeroPoint.Checked)
                {
                    _config.radioButtonZeroPoint = false;
                    _config.radioButtonNoZeroPoint = true;
                    Configuration.Save(_config);
                }
                    
            }
        }

        private void mainForm_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Move;
            }
        }

        private void mainForm_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files.Length == 1)
            {
                string Extension = Path.GetExtension(files[0]).ToLower();

                if (Extension == ".csv")
                {
                    string FilePath = files[0];
                    if (File.Exists(FilePath) && Path.GetExtension(FilePath).ToLower() == ".csv")
                    {
                        OpenPPFcsv(FilePath);
                    }
                }
            }
        }

        private void dataGridViewgetDXFPick_RowStateChanged(object sender, DataGridViewRowStateChangedEventArgs e)
        {
            e.Row.HeaderCell.Value = string.Format("{0}", e.Row.Index + 1);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            //CsvWriter csvWriter = new CsvWriter()
            SaveFileDialog saveFile = new SaveFileDialog();

            saveFile.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
            saveFile.FilterIndex = 1;
            if (saveFile.ShowDialog() == DialogResult.OK)
            {
                using (var fs = new FileStream(saveFile.FileName, FileMode.Create))
                using (var writer = new StreamWriter(fs, Encoding.Default))
                using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
                {
                    DataTable table = getDXFPickTableView.ToTable();


                    foreach (DataColumn column in table.Columns)
                    {
                        csv.WriteField(column.ColumnName);
                    }

                    csv.NextRecord();


                    foreach (DataRow row in table.Rows)
                    {
                        for (var i = 0; i < table.Columns.Count; i++)
                        {
                            csv.WriteField(row[i]);
                        }

                        csv.NextRecord();
                    }

                    csv.Flush();

                }
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (getDXFPickTable.Columns.Contains("NO.") == false)
            {
                getDXFPickTable.Columns.Add("NO.").SetOrdinal(0);
            }
            if (getDXFPickTable.Columns.Contains("Designator") == false)
            {
                getDXFPickTable.Columns.Add("Designator").SetOrdinal(1);
            }
            if (getDXFPickTable.Columns.Contains("Comment") == false)
            {
                getDXFPickTable.Columns.Add("Comment");
            }
            if (getDXFPickTable.Columns.Contains("Footprint") == false)
            {
                getDXFPickTable.Columns.Add("Footprint");
            }
            if (getDXFPickTable.Columns.Contains("Layer") == false)
            {
                getDXFPickTable.Columns.Add("Layer");
            }

            for (int i = 0; i < getDXFPickTable.Rows.Count; i++)
            {
                string NO = "";
                string Designator = "";
                string Comment = "";
                string Footprint = "";
                string Layer = "";

                string BlockName = "";
                string Description = "";

                BlockName = getDXFPickTable.Rows[i]["BlockName"].ToString();
                Description = getDXFPickTable.Rows[i]["Description"].ToString();

                string[] temp = BlockName.Split('#');
                if (temp.Length >= 1)
                {
                    Designator = temp[0];
                }
                if (temp.Length >= 2)
                {
                    NO =  temp[1];
                }
                if (temp.Length >= 3)
                {
                    Layer = temp[2];
                }
                getDXFPickTable.Rows[i]["NO."] = NO;
                getDXFPickTable.Rows[i]["Designator"] = Designator;
                getDXFPickTable.Rows[i]["Layer"] = Layer;

                int index = Description.IndexOf("]_[");

                
                if (index >= 0 )
                {
                    Comment = Description.Substring(1, index-1);
                    Footprint = Description.Substring(index+3, Description.Length - index -4);

                    getDXFPickTable.Rows[i]["Comment"] = Comment;
                    getDXFPickTable.Rows[i]["Footprint"] = Footprint;
                }

            }
            getDXFPickTableView = getDXFPickTable.DefaultView;
            dataGridViewgetDXFPick.DataSource = null;
            dataGridViewgetDXFPick.DataSource = getDXFPickTableView;
            button8.Enabled = false;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.Title = "EXCEL";
            openFileDialog1.Filter = "EXCEL (*xlsx;*xls)|*xlsx;*xls";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                openRules(openFileDialog1.FileName);
            }
        }

        void openRules(string FilePath)
        {
            DataTable dataTable = Excel.ToDataTable(FilePath, "Rules");
            if (dataTable == null)
            {
                MessageBox.Show("仅支持固定格式的文件, Excel 文件中必须有<Rules> 工作表");
                return;
            }

            //---- 删除空列 空行
            dataTable = RemoveEmptyRows(dataTable);
            dataTable = RemvoeEmptyColumns(dataTable);

            //---- 找必须存在的列
            string temp = String.Join( " ",ColumnName_dxfBlockRules);

            foreach (string ColumnName in ColumnName_dxfBlockRules)
            {
                //bool exist = false;

                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    string temp_i = dataTable.Columns[i].ColumnName.Trim().Replace(" ", "");
                    if (ColumnName == temp_i)
                    {
                        dataTable.Columns[i].ColumnName = ColumnName;
                        temp = temp.Replace(ColumnName,"");
                        break;
                    }
                }

            }
            if (string.IsNullOrEmpty(temp.Trim()) == false)
            {
                MessageBox.Show("打开的格式不正确,缺少列:" + temp);
                return;
            }

            if (dataTable.Rows.Count >0)
            {
                RulesTable.Clear();
            }

            for (int i = 0; i < dxfBlocksTable.Rows.Count; i++)
            {
                string BlockName = dxfBlocksTable.Rows[i]["Name"].ToString();
                bool HaveRules = false;
                for (int j = 0; j < dataTable.Rows.Count; j++)
                {
                    string excelBlockName = dataTable.Rows[j]["Block_Name"].ToString();
                    string Rules = dataTable.Rows[j]["匹配规则(正则表达式)"].ToString();
                    string RuleDescription = dataTable.Rows[j]["匹配规则说明"].ToString();
                    string CreationTime = dataTable.Rows[j]["规则创建时间"].ToString();
                    string Founder = dataTable.Rows[j]["规则创建人"].ToString();

                    if (BlockName == excelBlockName)
                    {
                        if (string.IsNullOrWhiteSpace(Rules) == false)
                        {
                            RulesTable.Rows.Add(BlockName, Rules, RuleDescription, CreationTime, Founder);
                            if (string.IsNullOrWhiteSpace(dxfBlocksTable.Rows[i]["匹配规则(正则表达式)"].ToString()))
                            {
                                dxfBlocksTable.Rows[i]["匹配规则(正则表达式)"] =  Rules;
                            }
                            else
                            {
                                dxfBlocksTable.Rows[i]["匹配规则(正则表达式)"] += "  或  " + Rules;
                            }
                            
                            HaveRules = true;
                        } 
                    }
                }
                if (HaveRules == false)
                {
                    dxfBlocksTable.Rows[i]["匹配规则(正则表达式)"] = "";
                }
            }


            if (RulesTable.Rows.Count >0 )
            {
                textBoxBlocksRulesFilePath.Text = FilePath;


                string binPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                binPath = Path.GetDirectoryName(binPath);


                string RulesFilePath;

                if (binPath == Path.GetDirectoryName(FilePath))
                {
                    RulesFilePath = Path.GetFileName(FilePath);
                }
                else
                {
                    RulesFilePath = FilePath;
                }


                if (_config.textBoxBlocksRulesFilePath != RulesFilePath)
                {
                    _config.textBoxBlocksRulesFilePath = RulesFilePath;
                    Configuration.Save(_config);
                }
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            //CsvWriter csvWriter = new CsvWriter()
            SaveFileDialog saveFile = new SaveFileDialog();

            saveFile.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
            saveFile.FilterIndex = 1;
            if (saveFile.ShowDialog() == DialogResult.OK)
            {
                using (var fs = new FileStream(saveFile.FileName, FileMode.Create))
                using (var writer = new StreamWriter(fs, Encoding.Default))
                using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
                {
                    foreach (DataColumn column in PickandPlaceTable.Columns)
                    {
                        csv.WriteField(column.ColumnName);
                    }

                    csv.NextRecord();


                    foreach (DataRow row in PickandPlaceTable.Rows)
                    {
                        for (var i = 0; i < PickandPlaceTable.Columns.Count; i++)
                        {
                            csv.WriteField(row[i]);
                        }

                        csv.NextRecord();
                    }

                    csv.Flush();

                }
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            openFileDialog.Filter = "DXF files (*.dxf)|*.dxf";
            openFileDialog.FilterIndex = 1;
            openFileDialog.Multiselect = true;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                foreach (var FilePath in openFileDialog.FileNames)
                {
                    setTBaseMap_Dxf(FilePath);
                }
            }
        }

        private void setTBaseMap_Dxf(string FilePath)
        {
            try
            {
                DxfDocument doc = DxfDocument.Load(FilePath);

                TBaseMap_Dxf = doc;
                textBox1.Text = FilePath;

            }
            catch (Exception)
            {

                throw;
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            openFileDialog.Filter = "DXF files (*.dxf)|*.dxf";
            openFileDialog.FilterIndex = 1;
            openFileDialog.Multiselect = true;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                foreach (var FilePath in openFileDialog.FileNames)
                {
                    setBBaseMap_Dxf(FilePath);
                }
            }
        }

        private void setBBaseMap_Dxf(string FilePath)
        {
            try
            {
                DxfDocument doc = DxfDocument.Load(FilePath);

                BBaseMap_Dxf = doc;
                textBox2.Text = FilePath;

            }
            catch (Exception)
            {

                throw;
            }
        }

        private void checkBoxBLayerMirrorR_CheckedChanged(object sender, EventArgs e)
        {
            if (_config.checkBoxBLayerMirrorR != checkBoxBLayerMirrorR.Checked)
            {
                _config.checkBoxBLayerMirrorR = checkBoxBLayerMirrorR.Checked;
                Configuration.Save(_config);
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            //使用第一引脚计算象限,根据第一引脚所在象限算角度
            if (CheckPadXOK==false || CheckPadYOK == false)
            {
                MessageBox.Show("必须有第一引脚坐标才能计算角度,列名称一般是:\"Pad X  Pad Y\"");
                return;
            }
            
                
            if (PickandPlaceTable.Columns.Contains("FixRotation") == false)
            {
                PickandPlaceTable.Columns.Add("FixRotation");
            }
            //if (PickandPlaceTable.Columns.Contains("Direction") == false)
            //{
            //    PickandPlaceTable.Columns.Add("Direction");
            //}
            if (PickandPlaceTable.Columns.Contains("FixStatus") == false)
            {
                PickandPlaceTable.Columns.Add("FixStatus");
            }

            string Lay = "";
                
            double MidX = 0; 
            double MidY = 0;
            double PadX = 0;
            double PadY = 0;
            double Rotation = 0; //当前坐标提供的角度
            double FixRotation = 0;
            string isFix = "";

            double PadDirection = 0; 

            Vector2 PadXY = new Vector2();

            for (int i = 0; i < PickandPlaceTable.Rows.Count; i++)
            {
                string Designator = "";
                string Comment = "";
                string Footprint = "";

                Designator = PickandPlaceTable.Rows[i][ColumnName_Designator[0]].ToString().Trim();

                if (CheckCommentOK)
                {
                    Comment = PickandPlaceTable.Rows[i][ColumnName_Comment[0]].ToString().Trim();
                }
                if (CheckFootprintOK)
                {
                    Footprint = PickandPlaceTable.Rows[i][ColumnName_Footprint[0]].ToString().Trim();
                }

                if (string.IsNullOrEmpty(Footprint) == true)
                {
                    continue;
                }
                MidX = Convert.ToDouble(PickandPlaceTable.Rows[i][ColumnName_MidX[0]].ToString().Trim());
                MidY = Convert.ToDouble(PickandPlaceTable.Rows[i][ColumnName_MidY[0]].ToString().Trim());
                Rotation = Convert.ToDouble(PickandPlaceTable.Rows[i][ColumnName_Rotation[0]].ToString().Trim());
                Lay = PickandPlaceTable.Rows[i][ColumnName_Layer[0]].ToString().Trim();
                if (string.IsNullOrWhiteSpace(Lay))
                {
                    continue;
                }

                if (string.IsNullOrWhiteSpace(PickandPlaceTable.Rows[i][ColumnName_PadX[0]].ToString().Trim()) )
                {
                    continue;
                }
                else
                {
                    PadX = Convert.ToDouble(PickandPlaceTable.Rows[i][ColumnName_PadX[0]].ToString().Trim());
                }

                if (string.IsNullOrWhiteSpace(PickandPlaceTable.Rows[i][ColumnName_PadY[0]].ToString().Trim()))
                {
                    continue;
                }
                else
                {
                    PadY = Convert.ToDouble(PickandPlaceTable.Rows[i][ColumnName_PadY[0]].ToString().Trim());
                }

                if (checkBoxBLayerMirrorX.Checked && Lay == BLayerName[0]) //B层是否镜像?
                {
                    MidX = (MidX - MidX * 2);
                    PadX = (PadX - PadX * 2);
                }
                if (checkBoxBLayerMirrorR.Checked && Lay == BLayerName[0])  //B  层角度镜像?
                {
                    Rotation = 180 - Rotation;
                    Rotation = Rotation < 0 ? Rotation + 360 : Rotation;
                }

                PadXY = spinPoint(MidX, MidY, 360-Rotation,  PadX, PadY);  //旋转到0角度的情况

                double angleOfLine = Math.Atan2(( PadXY.Y - MidY  ), ( PadXY.X - MidX)) * 180 / Math.PI;  //(y,x)

                angleOfLine = angleOfLine < 0 ? angleOfLine + 360 : angleOfLine;

                PadDirection = angleOfLine;


                //查出是否有块
                DataView BlocksTableView = dxfBlocksTable.DefaultView;
                BlocksTableView.RowFilter = "Name = '" + Footprint + "'";
                DataTable BlocksTableFiltertable = BlocksTableView.ToTable();

                FixRotation = Rotation;
                isFix = "";

                if (BlocksTableFiltertable.Rows.Count > 0)
                {   //有完全相同的块名称

                    string BlockD = BlocksTableFiltertable.Rows[0]["Direction"].ToString();
                    if (string.IsNullOrWhiteSpace(BlockD) == false)
                    {
                        double BlockDirection = Convert.ToDouble(BlockD);

                        //t u v w
                        double u = PadDirection - BlockDirection ;
                        double v = u % 90;
                        int t = (int)(u / 90);
                        if (v == 0 && u != 0) //检查相差整数度
                        {
                            FixRotation += u;
                            FixRotation = (FixRotation+360) % 360;

                            isFix = "修正角度:" + u.ToString();
                        }
                        else if (Math.Abs(v) > 45) //角度差大于45,才考虑是否增加角度, 否则不用做旋转操作
                        {
                            if (v > 0) //增加方向, 反之减方向
                            {
                                int w = ((int)(u / 90)+1) * 90;
                                FixRotation += w;
                                isFix = "增加方向";
                            }
                            else
                            {
                                int w = ((int)(u / 90) - 1) * 90;
                                FixRotation += w;
                                isFix = "减方向";
                            }
                            FixRotation = FixRotation % 360;
                        }
                        else if(t != 0)
                        {
                            if (t > 0) //增加方向, 反之减方向
                            {
                                int w = t * 90;
                                FixRotation += w;
                                isFix = "增加方向";
                            }
                            else
                            {
                                int w = t* 90;
                                FixRotation += w;
                                isFix = "减方向";
                            }
                            FixRotation = FixRotation % 360;
                        }
                    }
                }
                else if (radioButtonRulesEnabled.Checked)
                {
                    bool fixOK = false;
                    foreach (DataRow dxfBlockRow in dxfBlocksTable.Rows)
                    {
                        DataView RulesView = RulesTable.DefaultView;
                        RulesView.RowFilter = "Block_Name = '" + dxfBlockRow["Name"] + "'";
                        DataTable RowFiltertable = RulesView.ToTable();

                        for (int r = 0; r < RowFiltertable.Rows.Count; r++)
                        {
                            string Rule = RowFiltertable.Rows[r]["匹配规则(正则表达式)"].ToString();
                            if (Regex.IsMatch(Footprint, Rule))
                            {
                                fixOK = true;
                                string BlockD = dxfBlockRow["Direction"].ToString();
                                if (string.IsNullOrWhiteSpace(BlockD) == false)
                                {
                                    double BlockDirection = Convert.ToDouble(BlockD);

                                    //u v w
                                    double u = PadDirection - BlockDirection;
                                    double v = u % 90;
                                    int t = (int)(u / 90);
                                    if (v == 0 && u != 0) //检查相差整数度
                                    {
                                        FixRotation += u;
                                        FixRotation = (FixRotation + 360) % 360;

                                        isFix = "修正角度:" + u.ToString(); 
                                    }
                                    else if (Math.Abs(v) > 45) //角度差大于45,才考虑是否增加角度, 否则不用做旋转操作
                                    {
                                        if (v > 0) //增加方向, 反之减方向
                                        {
                                            int w = ((int)(u / 90) + 1) * 90;
                                            FixRotation += w;
                                            isFix = "增加方向";
                                        }
                                        else
                                        {
                                            int w = ((int)(u / 90) - 1) * 90;
                                            FixRotation += w;
                                            isFix = "减方向";
                                        }
                                        FixRotation = (FixRotation + 360) % 360;
                                    }
                                    else if (t != 0)
                                    {
                                        if (t > 0) //增加方向, 反之减方向
                                        {
                                            int w = t * 90;
                                            FixRotation += w;
                                            isFix = "增加方向";
                                        }
                                        else
                                        {
                                            int w = t * 90;
                                            FixRotation += w;
                                            isFix = "减方向";
                                        }
                                        FixRotation = FixRotation % 360;
                                    }

                                    break;
                                }
                            }
                        }
                    }

                    if (fixOK== false)
                    {
                        isFix = "按照规则搜索Block 库,也没有同名封装,暂无法修正";
                    }
                }
                else
                {
                    isFix = "Block 中没有同名封装,暂无法修正";
                }

                //B层角度还要还原回去
                if (checkBoxBLayerMirrorR.Checked && Lay == BLayerName[0])  //B  层角度镜像?
                {
                    FixRotation = (FixRotation + 180);
                    FixRotation = FixRotation % 360;
                    //顺时针角度

                    //转是逆时针角度
                    FixRotation = 360 - FixRotation;
                }

                PickandPlaceTable.Rows[i]["FixRotation"] = FixRotation.ToString();
                //PickandPlaceTable.Rows[i]["Direction"] = PadDirection.ToString();
                PickandPlaceTable.Rows[i]["FixStatus"] = isFix;
                BlocksTableView.RowFilter  = String.Empty;
            }

            MessageBox.Show("修正角度完成");
            checkBoxUsinFixRotation.Enabled = true;



        }
    }

    public class Excel
    {
        public static DataTable ToDataTable(string filePath, string sheetName)
        {
            DataTable dataTable = null;
            try
            {
                IWorkbook workbook;
                using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    workbook = WorkbookFactory.Create(stream);
                }
                ISheet sheet;
                if (string.IsNullOrWhiteSpace(sheetName))
                {
                     sheet = workbook.GetSheetAt(0); // zero-based index of your target sheet
                }
                else
                {
                     sheet = workbook.GetSheet(sheetName); // zero-based index of your target sheet
                }
                
                if (sheet == null)
                {
                    throw new Exception("不存在要读取的工作表:" + sheetName);
                }
                dataTable = new DataTable(sheet.SheetName);

                int rfirst = sheet.FirstRowNum;  //首行起始
                int rlast = sheet.LastRowNum;    //尾行
                IRow row = sheet.GetRow(rfirst); //获取首行
                int cfirst = row.FirstCellNum;   //首列起始
                int clast = row.LastCellNum;     //尾列

                for (int i = cfirst; i < clast; i++)
                {
                    if (row.GetCell(i) != null)
                        dataTable.Columns.Add(row.GetCell(i).StringCellValue, System.Type.GetType("System.String"));
                }
                row = null;
                for (int i = rfirst + 1; i <= rlast; i++)
                {
                    DataRow r = dataTable.NewRow();
                    IRow ir = sheet.GetRow(i);
                    if (ir != null)
                    {
                        for (int j = cfirst; j < clast; j++)
                        {
                            ICell cell = ir.GetCell(j);


                            if (cell != null)
                            {
                                //r[j] = ir.GetCell(j).ToString();
                                String cellValue = "";

                                switch (cell.CellType)
                                {
                                    case CellType.String:
                                        cellValue = cell.StringCellValue;
                                        r[j] = cellValue;
                                        break;
                                    case CellType.Numeric:
                                        cellValue = cell.NumericCellValue.ToString();
                                        r[j] = cellValue;
                                        break;
                                    case CellType.Formula:
                                        cellValue = cell.NumericCellValue.ToString();
                                        r[j] = cellValue;
                                        break;
                                    default:
                                        r[j] = "";
                                        break;

                                }
                            }
                            else
                            {
                                r[j] = "";
                            }
                        }
                        dataTable.Rows.Add(r);
                    }

                    ir = null;
                    r = null;
                }
                sheet = null;
                workbook = null;
            }
            catch (Exception e)
            {

                throw e;
            }

            return dataTable;
        }


        public static void SaveToExcel(DataTable dataTable, string sheetName, string filePath)
        {
            FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            IWorkbook workBook = null;

            if (Path.GetExtension(filePath).ToLower() == ".xlsx")
            {
                workBook = new XSSFWorkbook();
            }
            else
            {
                workBook = new HSSFWorkbook();
            }

            ISheet sheet = workBook.CreateSheet(string.IsNullOrWhiteSpace(sheetName) ? "sheet1" : sheetName);
            IRow row = sheet.CreateRow(0);
            //处理表格列头
            row = sheet.CreateRow(0);
            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                row.CreateCell(i).SetCellValue(dataTable.Columns[i].ColumnName);
                sheet.AutoSizeColumn(i);
            }

            //处理数据内容
            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                row = sheet.CreateRow(1 + i);
                row.Height = 250;
                for (int j = 0; j < dataTable.Columns.Count; j++)
                {
                    row.CreateCell(j).SetCellValue(dataTable.Rows[i][j].ToString());
                }
            }
            //写入数据流
            workBook.Write(fs);
            fs.Flush();
            fs.Close();

        }
    }

    public class Configuration
    {
        public static string CONFIG_FILE = "gui-config.json";

        public bool checkBoxTLayer;
        public bool checkBoxBLayer;
        public bool checkBoxBLayerMirrorX;
        public bool checkBoxBLayerMirrorR;

        public bool checkBoxDesignatorTXT;
        public bool checkBoxCommentTXT;
        public bool checkBoxFootprintTXT;

        public string textBoxTXTHeight;
        public bool checkBoxTXT180;

        public bool radioButtonNoDrawing;
        public bool radioButtonDrawing0402;
        public bool radioButtonDrawing0603;

        public bool radioButtonRulesEnabled;

        public bool radioButtonNoZeroPoint;
        public bool radioButtonZeroPoint;

        public string textBoxDXFBlocksFilePath;
        public string textBoxBlocksRulesFilePath;

        public Configuration()
        {
            
            textBoxDXFBlocksFilePath = @"PartsLibrary.dxf";
            textBoxBlocksRulesFilePath = "MatchingRulesTable.xlsx";

            textBoxTXTHeight = "0.8";


            checkBoxTLayer = true;
            checkBoxBLayer = true;
            checkBoxBLayerMirrorX = true;
            checkBoxBLayerMirrorR = true;

            checkBoxDesignatorTXT = true;
            checkBoxCommentTXT = false;
            checkBoxFootprintTXT = false;

            checkBoxTXT180 = true;

            radioButtonNoDrawing = false;
            radioButtonDrawing0402 = false;
            radioButtonDrawing0603 = true;

            radioButtonRulesEnabled = true;

            radioButtonNoZeroPoint = false;
            radioButtonZeroPoint = true;
        }

        public static Configuration Load()
        {
            return LoadFile(CONFIG_FILE);
        }

        public static Configuration LoadFile(string filename)
        {
            try
            {
                string configContent = File.ReadAllText(filename);

                return Load(configContent);
            }
            catch (Exception e)
            {
                if (!(e is FileNotFoundException))
                {
                    Console.WriteLine(e);
                }
                return new Configuration();
            }
        }


        public static Configuration Load(string config_str)
        {
            try
            {
                Configuration config = JsonConvert.DeserializeObject<Configuration>(config_str);
                return config;
            }
            catch
            {
            }
            return null;
        }
        public static void Save(Configuration config)
        {
            try
            {
                string str1 = System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData);

                if (Directory.Exists(str1 + "\\" + "Pick_and_Place_File_to_DXF") == false)//如果不存,在就创建file文件夹
                {
                    Directory.CreateDirectory(str1 + "\\" + "Pick_and_Place_File_to_DXF");
                }

                using (StreamWriter sw = new StreamWriter(File.Open(str1 + "\\" + "Pick_and_Place_File_to_DXF" + "\\" + CONFIG_FILE, FileMode.Create)))
                {
                    string jsonString = JsonConvert.SerializeObject(config);//SimpleJson.SimpleJson.SerializeObject(config);
                    sw.Write(jsonString);
                    sw.Flush();
                }
            }
            catch (IOException e)
            {
                Console.Error.WriteLine(e);
            }
        }
        public void CopyFrom(Configuration config)
        {
            checkBoxTLayer = config.checkBoxTLayer;
            checkBoxBLayer = config.checkBoxBLayer;
            checkBoxBLayerMirrorX = config.checkBoxBLayerMirrorX;
            checkBoxBLayerMirrorR = config.checkBoxBLayerMirrorR;

            checkBoxDesignatorTXT = config.checkBoxDesignatorTXT;
            checkBoxCommentTXT = config.checkBoxCommentTXT;
            checkBoxFootprintTXT = config.checkBoxFootprintTXT;

            textBoxTXTHeight = config.textBoxTXTHeight;
            checkBoxTXT180 = config.checkBoxTXT180;

            radioButtonNoDrawing = config.radioButtonNoDrawing;
            radioButtonDrawing0402 = config.radioButtonDrawing0402;
            radioButtonDrawing0603 = config.radioButtonDrawing0603;

            radioButtonRulesEnabled = config.radioButtonRulesEnabled;

            radioButtonNoZeroPoint = config.radioButtonNoZeroPoint;
            radioButtonZeroPoint = config.radioButtonZeroPoint;

            textBoxDXFBlocksFilePath = config.textBoxDXFBlocksFilePath;
            textBoxBlocksRulesFilePath = config.textBoxBlocksRulesFilePath;
        }

        public void FixConfiguration()
        {

        }

    }

}
