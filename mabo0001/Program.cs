using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using mabo0001;
using System.Data;


namespace mabo0001
{
    class Program
    {
        enum Gender
        {
            男, 女
        }

        static void Main(string[] args)
        {

            DicRelation dr = new DicRelation();
            
            MyWord myword = new MyWord("F:\\2.dot");
            
            DbAccess dbaccess = new DbAccess("f:\\2.mdb");
            string sqlCBF= "select * from CBF";
            string sqlCBDKDC = "select * from CBDKDC";
            string sqlCBF_JTCY = "select * from CBF_JTCY";
            string sqlCBDKXX = "select * from CBDKXX";
            System.Data.DataTable dataCBF = DbAccess.GetDataSet(sqlCBF);
            System.Data.DataRow[] dataCBF_JTCY = null;
            System.Data.DataRow[] dataCBDKDC = null;
            System.Data.DataRow[] dataCBDKXX = null;
            foreach (DataRow row in dataCBF.Rows)
            {
                myword.Open();
                //在书签处插入值
                myword.InsertValue("县", "闻喜");
                myword.InsertValue("乡","郭家庄");
                myword.InsertValue("村", "崔庄");
                myword.InsertValue("组", String.Format("{00}",row["CBFBM"].ToString().Substring(12,2)));

                Console.WriteLine(row["CBFBM"]);
                dataCBDKDC = DbAccess.GetDataSet(sqlCBDKDC).Select("CBFBM=" + row["CBFBM"]);
                dataCBF_JTCY = DbAccess.GetDataSet(sqlCBF_JTCY).Select("CBFBM=" + row["CBFBM"]);
                dataCBDKXX = DbAccess.GetDataSet(sqlCBDKXX).Select("CBFBM=" + row["CBFBM"]);
                Word.Table table = myword.doc.Tables[1];

                table.Cell(4, 1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                table.Cell(4, 1).Range.Text = row["CBFBM"].ToString().Substring(14);
                table.Cell(4, 1).Range.Font.Size = 10;
                table.Cell(4, 1).Range.Font.Spacing = float.Parse("0");

                table.Cell(4, 2).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                table.Cell(4, 2).Range.Text = row["CBFMC"].ToString();
                table.Cell(4, 2).Range.Font.Size = 10;
                table.Cell(4, 2).Range.Font.Spacing = float.Parse("0");
                

                table.Cell(4, 3).Range.Text = dataCBF_JTCY.Count().ToString();
                table.Cell(4, 3).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                table.Cell(4, 3).Range.Font.Size = 10;
                table.Cell(4, 3).Range.Font.Spacing = float.Parse("0");

                table.Cell(4, 11).Range.Text = dataCBDKDC.Count().ToString();
                table.Cell(4, 11).Range.Font.Size = 10;
                table.Cell(4, 11).Range.Font.Spacing = float.Parse("0");
                table.Cell(4, 11).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                table.Cell(4, 10).Range.Text = dataCBDKXX.Count()== 0 ? "/" : dataCBDKXX[0]["CBHTBM"].ToString();
                table.Cell(4, 10).Range.Font.Size = 10;
                table.Cell(4, 10).Range.Font.Spacing = float.Parse("0");
                table.Cell(4, 10).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;


                double HTarea = 0;
                double SCarea = 0;
                //table.Cell(4, 14).Range.Text = dataCBDKDC
                
                for (int i=0;i< dataCBF_JTCY.Count(); i++)
                {
                    table.Cell(4+i, 4).Range.Text = dataCBF_JTCY[i]["CYXM"].ToString();
                    table.Cell(4 + i, 4).Range.Font.Size = 10;
                    table.Cell(4 + i, 4).Range.Font.Spacing = float.Parse("0");
                    table.Cell(4 + i, 5).Range.Text = dataCBF_JTCY[i]["CYXB"].ToString()=="1" ? "男":"女";
                    table.Cell(4 + i, 5).Range.Font.Size = 10;
                    table.Cell(4 + i, 5).Range.Font.Spacing = float.Parse("0");

                    System.DateTime csrq = (System.DateTime)dataCBF_JTCY[i]["CSRQ"];
                    table.Cell(4 + i, 6).Range.Text = (System.DateTime.Now.Year - csrq.Year) < 0 ? "/" : (System.DateTime.Now.Year - csrq.Year).ToString();
                    table.Cell(4 + i, 6).Range.Font.Size = 10;
                    table.Cell(4 + i, 6).Range.Font.Spacing = float.Parse("0");

                    table.Cell(4 + i, 7).Range.Text = DicRelation.yCBFGX[dataCBF_JTCY[i]["YHZGX"].ToString()];
                    table.Cell(4 + i, 7).Range.Font.Size = 10;
                    table.Cell(4 + i, 7).Range.Font.Spacing = float.Parse("0");

                    table.Cell(4 + i, 8).Range.Text = dataCBF_JTCY[i]["SFGYR"].ToString() == "1" ? "是" : "否";
                    table.Cell(4 + i, 8).Range.Font.Size = 10;
                    table.Cell(4 + i, 8).Range.Font.Spacing = float.Parse("0");

                    table.Cell(4 + i, 9).Range.Text = DicRelation.cYBZ[dataCBF_JTCY[i]["CYBZ"].ToString()];
                    table.Cell(4 + i, 9).Range.Font.Size = 10;
                    table.Cell(4 + i, 9).Range.Font.Spacing = float.Parse("0");

                }

                for (int j =0;j<dataCBDKDC.Count(); j++)
                {
                    HTarea = HTarea + (double)dataCBDKDC[j]["HTMJ"];
                    SCarea = SCarea + (double)dataCBDKDC[j]["SCMJ"]/666.66;
                    table.Cell(4 + j, 12).Range.Text = dataCBDKDC[j]["DKMC"].ToString();
                    table.Cell(4 + j, 12).Range.Font.Size = 10;
                    table.Cell(4 + j, 12).Range.Font.Spacing = float.Parse("0");

                    table.Cell(4 + j, 13).Range.Text = dataCBDKDC[j]["HTMJ"].ToString();
                    table.Cell(4 + j, 13).Range.Font.Size = 10;
                    table.Cell(4 + j, 13).Range.Font.Spacing = float.Parse("0");
                    table.Cell(4 + j, 15).Range.Text = dataCBDKDC[j]["DKDZ"].ToString().Split('临')[1];
                    table.Cell(4 + j, 15).Range.Font.Size = 10;
                    table.Cell(4 + j, 15).Range.Font.Spacing = float.Parse("0");
                    table.Cell(4 + j, 16).Range.Text = dataCBDKDC[j]["DKNZ"].ToString().Split('临')[1];
                    table.Cell(4+  j, 16).Range.Font.Size = 10;
                    table.Cell(4 + j, 16).Range.Font.Spacing = float.Parse("0");
                    table.Cell(4 + j, 17).Range.Text = dataCBDKDC[j]["DKXZ"].ToString().Split('临')[1];
                    table.Cell(4 + j, 17).Range.Font.Size = 10;
                    table.Cell(4 + j, 17).Range.Font.Spacing = float.Parse("0");
                    table.Cell(4 + j, 18).Range.Text = dataCBDKDC[j]["DKBZ"].ToString().Split('临')[1];
                    table.Cell(4 + j, 18).Range.Font.Size = 10;
                    table.Cell(4 + j, 18).Range.Font.Spacing = float.Parse("0");
                    table.Cell(4 + j, 20).Range.Text = String.Format("{0:0.00}",(double)dataCBDKDC[j]["SCMJ"]/666.67);
                    table.Cell(4 + j, 20).Range.Font.Size = 10;
                    table.Cell(4 + j, 20).Range.Font.Spacing = float.Parse("0");
                }
                
                table.Cell(4, 14).Range.Text =  HTarea.ToString();
                table.Cell(4, 14).Range.Font.Size = 10;
                table.Cell(4, 14).Range.Font.Spacing = float.Parse("0");
                table.Cell(4, 14).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                Console.WriteLine("1");
                myword.SaveDocument("F:\\1234\\" + row["CBFBM"] + ".doc");
            }
            
        }
    }

}
