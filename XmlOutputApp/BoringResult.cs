using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WorldGyomu.CommonControl; //待機ダイアログ表示
using Microsoft.Office.Interop.Excel; //Excel出力処理に必要 2021/6/22 a_nagata↓↓↓↓
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel; //Excel出力処理に必要 2021/6/22 a_nagata↑↑↑↑
using System.Xml.Linq; //追加  Xmlを読み込むのに必要 2021/6/23 a_nagata

namespace XmlOutputApp
{
	public partial class BoringResult : Form
	{
		public BoringResult()
		{
			InitializeComponent();
		}

		/// <summary>
		/// セルダブルクリック時処理
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		// <param name="ws1"></param>  //イベントの引数変更×
		private void btnGanban_Click(object sender, EventArgs e) //, Excel.Worksheet ws1  イベントの引数変更×
		{
			#region "work 貼り付けるXMLを作成"

			//ここに配列として岩盤ボーリングのXMLを取り込む

			// 10行4列イメージの二次元配列
			//object[,] setValue = new object[10, 4];

			//	for (int i = 0; i < 10; i++)
			//	{
			//		for (int j = 0; j < 4; j++)
			//		{
			//			setValue[i, j] = string.Format("{0}行{1}列目", i + 1, j + 1);
			//		}
			//	}

			#endregion

			// Excel操作用オブジェクト
			Microsoft.Office.Interop.Excel.Application xlApp = null;
			Microsoft.Office.Interop.Excel.Workbook xlBook = null;
			Microsoft.Office.Interop.Excel.Worksheet xlSheet = null;

			String fName = "岩盤ボーリング柱状図" + ".xlsx";

			string spf = System.Windows.Forms.Application.StartupPath + "\\ExcelTemplate"; // exeのパス取得

			if (!File.Exists(Path.Combine(spf, fName)))  //テンプレートファイルのチェック
			{
				MessageBox.Show("テンプレートファイルが存在しません。");
				return;
			}

			// オープンファイルダイアログを生成する
			OpenFileDialog op = new OpenFileDialog();
			op.Title = "ファイルを開く";
			op.InitialDirectory = @"C:\";
			op.FileName = @"";
			op.Filter = "xmlファイル(*.xml;*)|*.xml;* | すべてのファイル(*.*)|*.*";
			op.FilterIndex = 1;

			//待機状態のマウス・カーソルを表示
			Cursor preCursor = Cursor.Current;

			//オープンファイルダイアログを表示する
			DialogResult result = op.ShowDialog();

			//拡張子判別
			string fileExt = System.IO.Path.GetExtension(op.FileName);

			if (result == DialogResult.OK)
			{
				//「開く」ボタンが選択された時の処理
				string fileName = op.FileName;  //ファイルパス取得
			}
			else if (result == DialogResult.Cancel)
			{
				//「キャンセル」ボタンまたは「×」ボタンが選択された時の処理
				return;
			}

			// マウスカーソルを待機状態にする
			Cursor.Current = Cursors.WaitCursor;
			// 待機ダイアログ表示
			WaitSplash.ShowSplash();

			// Excelアプリケーション生成
			xlApp = new Microsoft.Office.Interop.Excel.Application();

			// Excel表示
			xlApp.Visible = true;

			// エクセル保存確認表示
			xlApp.DisplayAlerts = true;

			// ◆新規のExcelブックを開く◆
			xlBook = xlApp.Workbooks.Add(spf + "\\" + fName);
			xlSheet = (Excel.Worksheet)xlBook.Worksheets[1];

			//罫線 外枠の範囲指定
			Excel.Range range_line;
			range_line = xlSheet.Range[xlSheet.Cells[2, 1], xlSheet.Cells[50, 20]];

			Excel.Borders borders;  //型宣言
			Excel.Border border;    //型宣言
			//罫線の設定
			borders = range_line.Borders;
			border = borders[Excel.XlBordersIndex.xlEdgeLeft];
			border.LineStyle = Excel.XlLineStyle.xlContinuous;
			border = borders[Excel.XlBordersIndex.xlEdgeRight];
			border.LineStyle = Excel.XlLineStyle.xlContinuous;
			border = borders[Excel.XlBordersIndex.xlEdgeTop];
			border.LineStyle = Excel.XlLineStyle.xlContinuous;
			border = borders[Excel.XlBordersIndex.xlEdgeBottom];
			border.LineStyle = Excel.XlLineStyle.xlContinuous;

			//内枠
			border = borders[Excel.XlBordersIndex.xlInsideVertical];
			border.LineStyle = Excel.XlLineStyle.xlContinuous;
			border = borders[Excel.XlBordersIndex.xlInsideHorizontal];
			border.LineStyle = Excel.XlLineStyle.xlContinuous;

			// XMLファイルの読み込み
			XElement xml = XElement.Load(op.FileName);

			//それぞれのタグ内の情報を取得する
			IEnumerable<String> borings = from item in xml.Elements("標題情報").Elements("調査基本情報").Elements("ボーリング名")
										select item.Value;

			IEnumerable<String> hyoukous = from item in xml.Elements("標題情報").Elements("ボーリング基本情報").Elements("孔口標高")
										select item.Value;

			IEnumerable<String> sousakukouchous = from item in xml.Elements("標題情報").Elements("ボーリング基本情報").Elements("総削孔長")
										select item.Value;

			//全項目の深度抽出
			IEnumerable<String> sindos = from item in xml.Elements("コア情報").Elements("工学的地質区分名現場土質名").Elements("工学的地質区分名現場土質名_下端深度")
										select item.Value;

			IEnumerable<String> sindo2s = from item in xml.Elements("コア情報").Elements("色調").Elements("色調_下端深度")
										select item.Value;

			IEnumerable<String> sindo3s = from item in xml.Elements("コア情報").Elements("風化の程度区分").Elements("風化の程度区分_下端深度")
										select item.Value;

			IEnumerable<String> sindo4s = from item in xml.Elements("コア情報").Elements("熱水変質の程度区分").Elements("熱水変質の程度区分_下端深度")
										select item.Value;

			IEnumerable<String> sindo5s = from item in xml.Elements("コア情報").Elements("硬軟区分").Elements("硬軟区分_下端深度")
										select item.Value;

			IEnumerable<String> sindo6s = from item in xml.Elements("コア情報").Elements("ボーリングコアの形状区分").Elements("ボーリングコアの形状区分_下端深度")
										select item.Value;

			IEnumerable<String> sindo7s = from item in xml.Elements("コア情報").Elements("割れ目の状態区分").Elements("割れ目の状態区分_下端深度")
										select item.Value;

			IEnumerable<String> sindo8s = from item in xml.Elements("コア情報").Elements("岩級区分").Elements("岩級区分_下端深度")
										select item.Value;

			IEnumerable<String> sindo9s = from item in xml.Elements("コア情報").Elements("コア採取率").Elements("コア採取率_下端深度")
										select item.Value;

			IEnumerable<String> sindo10s = from item in xml.Elements("コア情報").Elements("最大コア長").Elements("最大コア長_下端深度")
										select item.Value;

			IEnumerable<String> sindo11s = from item in xml.Elements("コア情報").Elements("RQD").Elements("RQD_下端深度")
										select item.Value;
			//全項目の深度抽出

			IEnumerable<String> dositumeis = from item in xml.Elements("コア情報").Elements("工学的地質区分名現場土質名").Elements("工学的地質区分名現場土質名_工学的地質区分名現場土質名")
										select item.Value;

			IEnumerable<String> sikichous = from item in xml.Elements("コア情報").Elements("色調").Elements("色調_色調名")
										select item.Value;

			IEnumerable<String> fukas = from item in xml.Elements("コア情報").Elements("風化の程度区分判定表").Elements("風化の程度区分判定表_記号")
										select item.Value;

			IEnumerable<String> nessuis = from item in xml.Elements("コア情報").Elements("熱水変質の程度区分判定表").Elements("熱水変質の程度区分判定表_記号")
										select item.Value;

			IEnumerable<String> kounans = from item in xml.Elements("コア情報").Elements("硬軟区分判定表").Elements("硬軟区分判定表_記号")
										select item.Value;

			IEnumerable<String> boring_cores = from item in xml.Elements("コア情報").Elements("ボーリングコアの形状区分判定表").Elements("ボーリングコアの形状区分判定表_記号")
										select item.Value;

			IEnumerable<String> waremes = from item in xml.Elements("コア情報").Elements("割れ目の状態区分判定表").Elements("割れ目の状態区分判定表_記号")
										select item.Value;

			IEnumerable<String> iwakyus = from item in xml.Elements("コア情報").Elements("岩級区分判定表").Elements("岩級区分判定表_判定").Elements("岩級区分判定表_記号")
										select item.Value;

			IEnumerable<String> saisyuritus = from item in xml.Elements("コア情報").Elements("コア採取率").Elements("コア採取率_採取率")
										select item.Value;

			IEnumerable<String> core_chous = from item in xml.Elements("コア情報").Elements("最大コア長").Elements("最大コア長_コア長")
										select item.Value;

			IEnumerable<String> rqds = from item in xml.Elements("コア情報").Elements("RQD").Elements("RQD_RQD")
										select item.Value;

			IEnumerable<String> kizis = from item in xml.Elements("コア情報").Elements("観察記事").Elements("観察記事_記事")
										select item.Value;

			var konai_suiixs = (
				from p in xml.Descendants("孔内水位")
				where (p.Element("孔内水位_削孔状況コード").Value) == "1" || (p.Element("孔内水位_削孔状況コード").Value) == "9"
				orderby Decimal.Parse(p.Element("孔内水位_孔内水位").Value), (p.Element("孔内水位_測定年月日"))
				select p);

			IEnumerable<String> konai_suiis = from item in xml.Elements("コア情報").Elements("孔内水位").Elements("孔内水位_測定年月日")
										select item.Value;

			IEnumerable<String> konai_suii2s = from item in xml.Elements("コア情報").Elements("孔内水位").Elements("孔内水位_削孔状況コード")
										select item.Value;

			IEnumerable<String> konai_suii3s = from item in xml.Elements("コア情報").Elements("孔内水位").Elements("孔内水位_孔内水位")
										select item.Value;

			// 貼り付け位置
			int row = 3;
			int col = 1;

			//IEnumerable(孔内水位)を配列に変換
			string[] hyoukou_array = hyoukous.ToArray();

			string[] konai_array = konai_suiis.ToArray();
			string[] konai_array2 = konai_suii2s.ToArray();
			string[] konai_array3 = konai_suii3s.ToArray();

			//リスト作成
			List<string> konai_arrayList = new List<string>();
			List<string> konai_array2List = new List<string>();
			List<string> konai_array3List = new List<string>();
			List<string> konai_array_allList = new List<string>();
			List<string> array = new List<string>();
			List<String> tmp = new List<String>();
			List<String> sindoList = new List<String>();
			List<String> suiiList = new List<String>();

			//リストに追加
			foreach (string konai_suii in konai_array)
			{
				konai_arrayList.Add(konai_suii);  //年月日
			}

			foreach (string konai_suii2 in konai_array2)
			{
				konai_array2List.Add(konai_suii2);  //コード
			}

			foreach (string konai_suii3 in konai_array3)
			{
				konai_array3List.Add(konai_suii3);  //水位
			}


			// 深度を深度リストに格納  A
			foreach (String sindo in sindos)
			{
				sindoList.Add(sindo);
			}

			// 色調の深度を深度リストに格納  B
			for (int i = 0; i < sindoList.Count; i++)
			{
				foreach (String sindosiki in sindo2s)
				{
					if (sindoList[i].ToString() == sindosiki) // 同じ深度があれば次
						continue;

					if (sindoList.Count - 1 == i && Decimal.Parse(sindoList[i].ToString()) < Decimal.Parse(sindosiki))
						tmp.Add(sindosiki);

					else if (Decimal.Parse(sindoList[i].ToString()) < Decimal.Parse(sindosiki) && Decimal.Parse(sindosiki) < Decimal.Parse(sindoList[i + 1].ToString()))
						tmp.Add(sindosiki);
				}

				if (tmp.Count > 0)
				{
					sindoList.InsertRange(i + 1, tmp);
					tmp = new List<string>(); // tmp初期化
				}
			}

			//風化の深度を深度リストに格納  C
			for (int j = 0; j < sindoList.Count; j++)
			{
				foreach (String sindofu in sindo3s)
				{
					if (sindoList[j].ToString() == sindofu) //同じ深度があれば次
						continue;

					if (sindoList.Count - 1 == j && Decimal.Parse(sindoList[j].ToString()) < Decimal.Parse(sindofu))
						tmp.Add(sindofu);

					else if (Decimal.Parse(sindoList[j].ToString()) < Decimal.Parse(sindofu) && Decimal.Parse(sindofu) < Decimal.Parse(sindoList[j + 1].ToString()))
						tmp.Add(sindofu);
				}

				if (tmp.Count > 0)
				{
					sindoList.InsertRange(j + 1, tmp);
					tmp = new List<string>(); //tmp初期化
				}
			}

			//ボーリングコアの形状の深度を深度リストに格納  D
			for (int k = 0; k < sindoList.Count; k++)
			{
				foreach (String sindocore in sindo6s)
				{
					if (sindoList[k].ToString() == sindocore) //同じ深度があれば次
						continue;

					if (sindoList.Count - 1 == k && Decimal.Parse(sindoList[k].ToString()) < Decimal.Parse(sindocore))
						tmp.Add(sindocore);

					else if (Decimal.Parse(sindoList[k].ToString()) < Decimal.Parse(sindocore) && Decimal.Parse(sindocore) < Decimal.Parse(sindoList[k + 1].ToString()))
						tmp.Add(sindocore);
				}

				if (tmp.Count > 0)
				{
					sindoList.InsertRange(k + 1, tmp);
					tmp = new List<string>(); //tmp初期化
				}
			}

			foreach (String konai_suii3 in konai_suii3s)
			{
				suiiList.Add(konai_suii3);
			}

			// ボーリング名
			row = 3;  //初期表示位置
			foreach (String boring in borings)
			{
				xlSheet.Cells[row, col] = (boring);
				row++;
			}

			//孔口標高
			row = 3;  //初期表示位置
			foreach (String hyoukou in hyoukous)
			{
				xlSheet.Cells[row, col + 1] = (hyoukou);
				row++;
			}

			//総削孔長
			row = 3;  //初期表示位置
			foreach (String sousakukouchou in sousakukouchous)
			{
				xlSheet.Cells[row, col + 2] = (sousakukouchou);
				row++;
			}

			//標高
			row = 3;
			for (int h = 0, j = 0; h < sindoList.Count; h++)
			{
				xlSheet.Cells[row, col + 3] = Decimal.Parse(hyoukou_array[j]) - Decimal.Parse(sindoList[h]);
				row++;
			}

			//深度
			row = 3;
			for (int x = 0; x < sindoList.Count; x++)
			{
				xlSheet.Cells[row, col + 4] = (sindoList[x]);
				row++;
			}

			//上端深度～下端深度
			row = 3;
			for (int y = 0; y < sindoList.Count; y++)
			{
				if (row == 3)  //if (xlSheet.Cells[row, col + 6])
				{
					xlSheet.Cells[row, col + 5] = "0～" + sindoList[y];
					row++;
				}

				else if (row != 3)
				{
					xlSheet.Cells[row, col + 5] = sindoList[y - 1] + "～" + sindoList[y];
					//[y-1]は今より一つ前の要素の中身を表示 [y]は現在の繰り返し回数分の要素を表示
					row++;
				}
			}

			//層厚
			row = 3;
			for (int z = 0; z < sindoList.Count; z++)
			{
				if (row == 3)
				{
					xlSheet.Cells[row, col + 6] = sindoList[z];
					row++;
				}

				else if (row != 3)
				{
					xlSheet.Cells[row, col + 6] = Decimal.Parse(sindoList[z]) - Decimal.Parse(sindoList[z - 1]);
					row++;
				}
			}

			// 工学的地質区分名のデータ
			row = 3;  //初期表示位置
			foreach (String dositumei in dositumeis)
			{
				xlSheet.Cells[row, col + 7] = (dositumei);
				row++;
			}

			//色調のデータ
			row = 3;
			foreach (String sikichou in sikichous)
			{
				xlSheet.Cells[row, col + 8] = (sikichou);
				row++;
			}

			//風化の程度データ
			row = 3;
			foreach (String fuka in fukas)
			{
				xlSheet.Cells[row, col +9] = (fuka);
				row++;
			}

			//変質の程度データ
			row = 3;
			foreach (String nessui in nessuis)
			{
				xlSheet.Cells[row, col + 10] = (nessui);
				row++;
			}

			//硬軟データ
			row = 3;
			foreach (String kounan in kounans)
			{
				xlSheet.Cells[row, col + 11] = (kounan);
				row++;
			}

			//コア形状データ
			row = 3;
			foreach (String boring_core in boring_cores)
			{
				xlSheet.Cells[row, col + 12] = (boring_core);
				row++;
			}

			//割れ目の状態データ
			row = 3;
			foreach (String wareme in waremes)
			{
				xlSheet.Cells[row, col + 13] = (wareme);
				row++;
			}

			//岩級区分のデータ
			row = 3;
			foreach (String iwakyu in iwakyus)
			{
				xlSheet.Cells[row, col + 14] = (iwakyu);
				row++;
			}

			//コア採取率のデータ
			row = 3;
			foreach (String saisyuritu in saisyuritus)
			{
				xlSheet.Cells[row, col + 15] = (saisyuritu);
				row++;
			}

			//最大コア長のデータ
			row = 3;
			foreach (String core_chou in core_chous)
			{
				xlSheet.Cells[row, col + 16] = (core_chou);
				row++;
			}

			//rqdのデータ
			row = 3;
			foreach (String rqd in rqds)
			{
				xlSheet.Cells[row, col + 17] = (rqd);
				row++;
			}

			//記事のデータ
			row = 3;
			foreach (String kizi in kizis)
			{
				xlSheet.Cells[row, col + 18] = (kizi);
				row++;
			}

			try
			{
				if (xlSheet.Cells[row, col + 19] != null)
				{
					//孔内水位のデータ
					row = 3;
					foreach (var konai_suiix in konai_suiixs)
					{
						xlSheet.Cells[row, col + 19] = (konai_suiix.Element("孔内水位_測定年月日").Value) + "\n" + (konai_suiix.Element("孔内水位_孔内水位").Value);
						row++;
					}
				}
			}
			catch (Exception ie)  //例外が発生した時の処理
			{
				using (Form f = new Form())
				{
					f.TopMost = true; // フォームを常に最前面に表示する
					// 作成したフォームを親フォームとしてメッセージボックスに設定
					MessageBox.Show(f, "選択されたファイルは正しくありません"); // 結果、メッセージボックスも最前面に表示される
					return;
				}

				throw ie;
			}

			finally  //例外の有無にかかわらず、必ず最後に実行される処理
			{
				// 待機ダイアログ終了
				WaitSplash.CloseSplash();

				// マウスカーソルを元の状態に戻す
				Cursor.Current = preCursor;
			}
			
		}


		/// <summary>
		/// セルダブルクリック時処理
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		 // <param name="ws1"></param>  //イベントの引数変更×
		private void btnDositu_Click(object sender, EventArgs e)  //, Excel.Worksheet ws1
		{
			#region "work 貼り付けるXMLを作成"

			#endregion

			// Excel操作用オブジェクト
			Microsoft.Office.Interop.Excel.Application xlApp = null;
			Microsoft.Office.Interop.Excel.Workbook xlBook = null;
			Microsoft.Office.Interop.Excel.Worksheet xlSheet = null;

			// オープンファイルダイアログを生成する
			OpenFileDialog op = new OpenFileDialog();
			op.Title = "ファイルを開く";
			op.InitialDirectory = @"C:\";
			op.FileName = @"";
			op.Filter = "xmlファイル(*.xml;*)|*.xml;* | すべてのファイル(*.*)|*.*";
			op.FilterIndex = 1;

			//オープンファイルダイアログを表示する
			DialogResult result = op.ShowDialog();

			//拡張子判別
			string fileExt = System.IO.Path.GetExtension(op.FileName);

			if (result == DialogResult.OK)
			{
				//「開く」ボタンが選択された時の処理
				string fileName = op.FileName;  //ファイルパス取得
			}
			else if (result == DialogResult.Cancel)
			{
				//「キャンセル」ボタンまたは「×」ボタンが選択された時の処理
				return;
			}


			//待機状態のマウス・カーソルを表示
			Cursor preCursor = Cursor.Current;

			String fName = "土質ボーリング柱状図(標準貫入試験用)" + ".xlsx";

			string spf = System.Windows.Forms.Application.StartupPath + "\\ExcelTemplate"; // exeのパス取得

			if (!File.Exists(Path.Combine(spf, fName)))  //テンプレートファイルのチェック
			{
				MessageBox.Show("テンプレートファイルが存在しません。");
				return;
			}

			xlApp = new Microsoft.Office.Interop.Excel.Application(); // Excelアプリケーション生成
			xlApp.Visible = true; // 表示
			xlApp.DisplayAlerts = true; // エクセル保存確認表示
			// ◆新規のExcelブックを開く◆
			xlBook = xlApp.Workbooks.Add(spf + "\\" + fName);
			xlSheet = (Excel.Worksheet)xlBook.Worksheets[1];

			// マウスカーソルを待機状態にする
			Cursor.Current = Cursors.WaitCursor;

			// 待機ダイアログ表示
			WaitSplash.ShowSplash();

			//罫線
			Excel.Range range_line;
			range_line = xlSheet.Range[xlSheet.Cells[2, 1], xlSheet.Cells[80,14]];

			Excel.Borders borders;  //型宣言
			Excel.Border border;  //型宣言
			//罫線の設定
			borders = range_line.Borders;
			border = borders[Excel.XlBordersIndex.xlEdgeLeft];
			border.LineStyle = Excel.XlLineStyle.xlContinuous;
			border = borders[Excel.XlBordersIndex.xlEdgeRight];
			border.LineStyle = Excel.XlLineStyle.xlContinuous;
			border = borders[Excel.XlBordersIndex.xlEdgeTop];
			border.LineStyle = Excel.XlLineStyle.xlContinuous;
			border = borders[Excel.XlBordersIndex.xlEdgeBottom];
			border.LineStyle = Excel.XlLineStyle.xlContinuous;

			//内枠
			border = borders[Excel.XlBordersIndex.xlInsideVertical];
			border.LineStyle = Excel.XlLineStyle.xlContinuous;
			border = borders[Excel.XlBordersIndex.xlInsideHorizontal];
			border.LineStyle = Excel.XlLineStyle.xlContinuous;

			// XMLファイルの読み込み
			XElement xml = XElement.Load(op.FileName);

			//それぞれのタグ内の情報を取得する
			IEnumerable<String> borings = from item in xml.Elements("標題情報").Elements("調査基本情報").Elements("ボーリング名")
										select item.Value;

			IEnumerable<String> hyoukous = from item in xml.Elements("標題情報").Elements("ボーリング基本情報").Elements("孔口標高")
										select item.Value;

			IEnumerable<String> sousakukouchous = from item in xml.Elements("標題情報").Elements("ボーリング基本情報").Elements("総削孔長")
										select item.Value;

			//全項目の深度抽出
			IEnumerable<String> sindos = from item in xml.Elements("コア情報").Elements("工学的地質区分名現場土質名").Elements("工学的地質区分名現場土質名_下端深度")
										select item.Value;

			IEnumerable<String> sindo2s = from item in xml.Elements("コア情報").Elements("色調").Elements("色調_下端深度")
										select item.Value;

			IEnumerable<String> sindo3s = from item in xml.Elements("コア情報").Elements("相対密度稠度").Elements("相対密度稠度_下端深度")
										select item.Value;

			IEnumerable<String> sindo4s = from item in xml.Elements("コア情報").Elements("観察記事").Elements("観察記事_下端深度")
										select item.Value;

			IEnumerable<String> sindo5s = from item in xml.Elements("コア情報").Elements("孔内水位").Elements("孔内水位_孔内水位")
										select item.Value;

			IEnumerable<String> sindo6s = from item in xml.Elements("コア情報").Elements("標準貫入試験").Elements("標準貫入試験_開始深度")
										select item.Value;
			//全項目の深度抽出

			IEnumerable<String> dositumeis = from item in xml.Elements("コア情報").Elements("工学的地質区分名現場土質名").Elements("工学的地質区分名現場土質名_工学的地質区分名現場土質名")
										select item.Value;

			IEnumerable<String> sikichous = from item in xml.Elements("コア情報").Elements("色調").Elements("色調_色調名")
										select item.Value;

			IEnumerable<String> mitudos = from item in xml.Elements("コア情報").Elements("相対密度稠度").Elements("相対密度_コード")
										select item.Value;

			IEnumerable<String> mitudo2s = from item in xml.Elements("コア情報").Elements("相対密度稠度").Elements("相対密度_状態")
										select item.Value;

			IEnumerable<String> mitudo3s = from item in xml.Elements("コア情報").Elements("相対密度稠度").Elements("相対密度稠度_下端深度")
										   select item.Value;

			IEnumerable<String> chudos = from item in xml.Elements("コア情報").Elements("相対密度稠度").Elements("相対稠度_コード")
										select item.Value;

			IEnumerable<String> chudo2s = from item in xml.Elements("コア情報").Elements("相対密度稠度").Elements("相対稠度_状態")
										select item.Value;

			IEnumerable<String> kizis = from item in xml.Elements("コア情報").Elements("観察記事").Elements("観察記事_記事")
										select item.Value;

			IEnumerable<String> konai_suiis = from item in xml.Elements("コア情報").Elements("孔内水位").Elements("孔内水位_測定年月日")
										select item.Value;

			IEnumerable<String> konai_suii2s = from item in xml.Elements("コア情報").Elements("孔内水位").Elements("孔内水位_削孔状況コード")
										select item.Value;

			IEnumerable<String> konai_suii3s = from item in xml.Elements("コア情報").Elements("孔内水位").Elements("孔内水位_孔内水位")
										select item.Value;

			IEnumerable<String> hyozyuns = from item in xml.Elements("コア情報").Elements("標準貫入試験").Elements("標準貫入試験_合計打撃回数")
										select item.Value;

			var konai_suiixs = (
				from p in xml.Descendants("孔内水位")
				where (p.Element("孔内水位_削孔状況コード").Value) == "1" || (p.Element("孔内水位_削孔状況コード").Value) == "9"
				orderby Decimal.Parse(p.Element("孔内水位_孔内水位").Value), (p.Element("孔内水位_測定年月日"))
				select p);

			// 貼り付け位置
			int row = 3;
			int col = 1;

			//IEnumerableを配列に変換
			string[] hyoukou_array = hyoukous.ToArray();

			string[] konai_array = konai_suiis.ToArray();  //A'
			string[] konai_array2 = konai_suii2s.ToArray();  //B'
			string[] konai_array3 = konai_suii3s.ToArray();  //C'

			//リスト作成
			List<String> tmp = new List<String>();
			List<String> sindoList = new List<String>();
			List<String> suiiList = new List<String>();
			List<string> konai_arrayList = new List<string>();
			List<string> konai_array2List = new List<string>();
			List<string> konai_array3List = new List<string>();

			foreach (string konai_suii in konai_array)
			{
				konai_arrayList.Add(konai_suii);  //年月日
			}

			foreach (string konai_suii2 in konai_array2)
			{
				konai_array2List.Add(konai_suii2);  //コード
			}

			foreach (string konai_suii3 in konai_array3)
			{
				konai_array3List.Add(konai_suii3);  //水位

			}

			// 深度を深度リストに格納  A
			foreach (String sindo in sindos)
			{
				sindoList.Add(sindo);
			}

			// 色調の深度を深度リストに格納  B
			for (int i = 0; i < sindoList.Count; i++)
			{
				foreach (String sindosiki in sindo2s)
				{
					if (sindoList[i].ToString() == sindosiki) // 同じ深度があれば次
						continue;

					if (sindoList.Count - 1 == i && Decimal.Parse(sindoList[i].ToString()) < Decimal.Parse(sindosiki))
						tmp.Add(sindosiki);

					else if (Decimal.Parse(sindoList[i].ToString()) < Decimal.Parse(sindosiki) && Decimal.Parse(sindosiki) < Decimal.Parse(sindoList[i + 1].ToString()))
						tmp.Add(sindosiki);
				}

				if (tmp.Count > 0)
				{
					sindoList.InsertRange(i + 1, tmp);
					tmp = new List<string>(); // tmp初期化
				}
			}

			//相対密度を深度リストに格納 C
			for (int j = 0; j < sindoList.Count; j++)
			{
				foreach (String sindomitu in mitudo3s)
				{
					if (sindoList[j].ToString() == sindomitu)  //同じ深度があれば次
						continue;

					if (sindoList.Count - 1 == j && Decimal.Parse(sindoList[j].ToString()) < Decimal.Parse(sindomitu))
						tmp.Add(sindomitu);

					else if (Decimal.Parse(sindoList[j].ToString()) < Decimal.Parse(sindomitu) && Decimal.Parse(sindomitu) < Decimal.Parse(sindoList[j + 1].ToString()))
						tmp.Add(sindomitu);
				}

				if (tmp.Count > 0)
				{
					sindoList.InsertRange(j + 1, tmp);
					tmp = new List<string>();
				}
			}


			foreach (String konai_suii3 in konai_suii3s)
			{
				suiiList.Add(konai_suii3);
			}

			// ボーリング名
			row = 3;  //初期表示位置
			foreach (String boring in borings)
			{
				xlSheet.Cells[row, col] = (boring);
				row++;
			}

			//孔口標高
			row = 3;  //初期表示位置
			foreach (String hyoukou in hyoukous)
			{
				xlSheet.Cells[row, col + 1] = (hyoukou);
				row++;
			}

			//総削孔長
			row = 3;  //初期表示位置
			foreach (String sousakukouchou in sousakukouchous)
			{
				xlSheet.Cells[row, col + 2] = (sousakukouchou);
				row++;
			}

			//標高
			row = 3;  //初期表示位置
			for (int h = 0, j = 0; h < sindoList.Count; h++)
			{
				xlSheet.Cells[row, col + 3] = Decimal.Parse(hyoukou_array[j]) - Decimal.Parse(sindoList[h]);
				row++;
			}

			//深度
			row = 3;  //初期表示位置
			for (int x = 0; x < sindoList.Count; x++)
			{
				xlSheet.Cells[row, col + 4] = (sindoList[x]);
				row++;
			}

			//上端深度～下端深度
			row = 3;
			for (int y = 0; y < sindoList.Count; y++)
			{
				if (row == 3)  //if (xlSheet.Cells[row, col + 6])
				{
					xlSheet.Cells[row, col + 5] = "0～" + sindoList[y];
					row++;
				}

				else if (row != 3)
				{
					xlSheet.Cells[row, col + 5] = sindoList[y - 1] + "～" + sindoList[y];
					//[y-1]は今より一つ前の要素の中身を表示 [y]は現在の繰り返し回数分の要素を表示
					row++;
				}
			}

			//層厚
			row = 3;
			for (int z = 0; z < sindoList.Count; z++)
			{
				if (row == 3)
				{
					xlSheet.Cells[row, col + 6] = sindoList[z];
					row++;
				}

				else if (row != 3)
				{
					xlSheet.Cells[row, col + 6] = Decimal.Parse(sindoList[z]) - Decimal.Parse(sindoList[z - 1]);
					row++;
				}
			}

			// 工学的地質区分名のデータ
			row = 3;  //初期表示位置
			foreach (String dositumei in dositumeis)
			{
				xlSheet.Cells[row, col + 7] = (dositumei);
				row++;
			}

			//色調のデータ
			row = 3;  //初期表示位置
			foreach (String sikichou in sikichous)
			{
				xlSheet.Cells[row, col + 8] = (sikichou);
				row++;
			}

			//相対密度のデータ
			row = 3;  //初期表示位置
			foreach (String mitudo in mitudos)
			{
				xlSheet.Cells[row, col + 9] = (mitudo);
				row++;
			}

			//相対稠度のデータ
			row = 3;  //初期表示位置
			foreach (String chudo in chudos)
			{
				xlSheet.Cells[row, col + 10] = (chudo);
				row++;
			}

			//記事のデータ
			row = 3;  //初期表示位置
			foreach (String kizi in kizis)
			{
				xlSheet.Cells[row, col + 11] = (kizi);
				row++;
			}

			try  //例外が発生する可能性のある処理
			{
				if (xlSheet.Cells[row, col + 12] != null)
				{
					//孔内水位のデータ
					row = 3;
					foreach (var konai_suiix in konai_suiixs)
					{
						xlSheet.Cells[row, col + 12] = (konai_suiix.Element("孔内水位_測定年月日").Value) + "\n" + (konai_suiix.Element("孔内水位_孔内水位").Value);
						row++;
					}
				}
			}
			catch (Exception ie)  //例外が発生した時の処理
			{

				// 親フォームを作成
				using (Form f = new Form())
				{
					f.TopMost = true; // 親フォームを常に最前面に表示する
					// 作成したフォームを親フォームとしてメッセージボックスに設定
					MessageBox.Show(f, "選択されたファイルは正しくありません"); // 結果、メッセージボックスも最前面に表示される
					return;
				}

				throw ie;

			}

			finally  //例外の有無にかかわらず、必ず最後に実行される処理
			{
				// 待機ダイアログ終了
				WaitSplash.CloseSplash();

				// マウスカーソルを元の状態に戻す
				Cursor.Current = preCursor;
			}

			//標準貫入試験(N値)のデータ
			row = 3;  //初期表示位置
			foreach (String hyozyun in hyozyuns)
			{
				xlSheet.Cells[row, col + 13] = (hyozyun);
				row++;
			}
		}


		/// <summary>
		/// セルダブルクリック時処理
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		// <param name="ws1"></param>  イベントの引数変更×
		private void btnZisuberiAll_Click(object sender, EventArgs e)//, Excel.Worksheet ws1
		{
			#region "work 貼り付けるXMLを作成"

			#endregion

			// Excel操作用オブジェクト
			Microsoft.Office.Interop.Excel.Application xlApp = null;
			Microsoft.Office.Interop.Excel.Workbooks xlBooks = null;
			Microsoft.Office.Interop.Excel.Workbook xlBook = null;
			Microsoft.Office.Interop.Excel.Sheets xlSheets = null;
			Microsoft.Office.Interop.Excel.Worksheet xlSheet = null;

			// オープンファイルダイアログを生成する
			OpenFileDialog op = new OpenFileDialog();
			op.Title = "ファイルを開く";
			op.InitialDirectory = @"C:\";
			op.FileName = @"";
			op.Filter = "xmlファイル(*.xml;*)|*.xml;* | すべてのファイル(*.*)|*.*";
			op.FilterIndex = 1;

			//オープンファイルダイアログを表示する
			DialogResult result = op.ShowDialog();

			//拡張子判別
			string fileExt = System.IO.Path.GetExtension(op.FileName);

			if (result == DialogResult.OK)
			{
				//「開く」ボタンが選択された時の処理
				string fileName = op.FileName;  //ファイルパス取得
			}
			else if (result == DialogResult.Cancel)
			{
				//「キャンセル」ボタンまたは「×」ボタンが選択された時の処理
				return;
			}

			//待機状態のマウス・カーソルを表示
			Cursor preCursor = Cursor.Current;

			

			try  //例外が発生する可能性のある処理
			{
				String fName = "地すべりボーリング柱状図(オールコアボーリング用)" + ".xlsx";

				string spf = System.Windows.Forms.Application.StartupPath + "\\ExcelTemplate"; // exeのパス取得

				if (!File.Exists(Path.Combine(spf, fName)))  //テンプレートファイルのチェック
				{
					MessageBox.Show("テンプレートファイルが存在しません。");
					return;
				}
				xlApp = new Microsoft.Office.Interop.Excel.Application(); // Excelアプリケーション生成
				xlApp.Visible = true; // 表示
				xlApp.DisplayAlerts = true; // エクセル保存確認表示
				// ◆新規のExcelブックを開く◆
				xlBook = xlApp.Workbooks.Add(spf + "\\" + fName);
				xlSheet = (Excel.Worksheet)xlBook.Worksheets[1];

				// マウスカーソルを待機状態にする
				Cursor.Current = Cursors.WaitCursor;
				// 待機ダイアログ表示
				WaitSplash.ShowSplash();
			}
			catch (Exception ie)  //例外が発生した時の処理
			{
				// xlSheet解放
				if (xlSheet != null)
				{
					Marshal.ReleaseComObject(xlSheet);
					xlSheet = null;
				}

				// xlSheets解放
				if (xlSheets != null)
				{
					Marshal.ReleaseComObject(xlSheets);
					xlSheets = null;
				}

				// xlBook解放
				if (xlBook != null)
				{
					try
					{
						xlBook.Close();
					}
					finally
					{
						Marshal.ReleaseComObject(xlBook);
						xlBook = null;
					}
				}

				// xlBooks解放
				if (xlBooks != null)
				{
					Marshal.ReleaseComObject(xlBooks);
					xlBooks = null;
				}

				// xlApp解放
				if (xlApp != null)
				{
					try
					{
						// アラートを戻して終了
						xlApp.DisplayAlerts = true;
						xlApp.Quit();
					}
					finally
					{
						Marshal.ReleaseComObject(xlApp);
					}
				}

				throw ie;
			}

			finally  //例外の有無にかかわらず、必ず最後に実行される処理
			{
				// 待機ダイアログ終了
				WaitSplash.CloseSplash();

				// マウスカーソルを元の状態に戻す
				Cursor.Current = preCursor;
			}

			

			//罫線
			Excel.Range range_line;
			range_line = xlSheet.Range[xlSheet.Cells[2, 1], xlSheet.Cells[40, 16]];

			Excel.Borders borders;  //型宣言
			Excel.Border border;  //型宣言
			//罫線の設定
			borders = range_line.Borders;
			border = borders[Excel.XlBordersIndex.xlEdgeLeft];
			border.LineStyle = Excel.XlLineStyle.xlContinuous;
			border = borders[Excel.XlBordersIndex.xlEdgeRight];
			border.LineStyle = Excel.XlLineStyle.xlContinuous;
			border = borders[Excel.XlBordersIndex.xlEdgeTop];
			border.LineStyle = Excel.XlLineStyle.xlContinuous;
			border = borders[Excel.XlBordersIndex.xlEdgeBottom];
			border.LineStyle = Excel.XlLineStyle.xlContinuous;

			//内枠
			border = borders[Excel.XlBordersIndex.xlInsideVertical];
			border.LineStyle = Excel.XlLineStyle.xlContinuous;
			border = borders[Excel.XlBordersIndex.xlInsideHorizontal];
			border.LineStyle = Excel.XlLineStyle.xlContinuous;

			// XMLファイルの読み込み
			XElement xml = XElement.Load(op.FileName);

			//それぞれのタグ内の情報を取得する
			IEnumerable<String> borings = from item in xml.Elements("標題情報").Elements("調査基本情報").Elements("ボーリング名")
										select item.Value;

			IEnumerable<String> hyoukous = from item in xml.Elements("標題情報").Elements("ボーリング基本情報").Elements("孔口標高")
										select item.Value;

			IEnumerable<String> sousakukouchous = from item in xml.Elements("標題情報").Elements("ボーリング基本情報").Elements("総削孔長")
										select item.Value;

			IEnumerable<String> sindos = from item in xml.Elements("コア情報").Elements("工学的地質区分名現場土質名").Elements("工学的地質区分名現場土質名_下端深度")
										select item.Value;

			IEnumerable<String> sindo2s = from item in xml.Elements("コア情報").Elements("色調").Elements("色調_下端深度")
										select item.Value;

			IEnumerable<String> sindo3s = from item in xml.Elements("コア情報").Elements("風化の程度区分").Elements("風化の程度区分_下端深度")
										select item.Value;

			IEnumerable<String> sindo4s = from item in xml.Elements("コア情報").Elements("熱水変質の程度区分").Elements("熱水変質の程度区分_下端深度")
										select item.Value;

			IEnumerable<String> sindo5s = from item in xml.Elements("コア情報").Elements("破砕度").Elements("破砕度_下端深度")
										select item.Value;

			IEnumerable<String> sindo6s = from item in xml.Elements("コア情報").Elements("硬軟区分").Elements("硬軟区分_下端深度")
										select item.Value;

			IEnumerable<String> sindo7s = from item in xml.Elements("コア情報").Elements("コア採取率").Elements("コア採取率_下端深度")
										select item.Value;

			IEnumerable<String> sindo8s = from item in xml.Elements("コア情報").Elements("観察記事").Elements("観察記事_下端深度")
										select item.Value;

			IEnumerable<String> sindo9s = from item in xml.Elements("コア情報").Elements("コア質量").Elements("コア質量_下端深度")
										select item.Value;

			IEnumerable<String> dositumeis = from item in xml.Elements("コア情報").Elements("工学的地質区分名現場土質名").Elements("工学的地質区分名現場土質名_工学的地質区分名現場土質名")
										select item.Value;

			IEnumerable<String> sikichous = from item in xml.Elements("コア情報").Elements("色調").Elements("色調_色調名")
										select item.Value;

			IEnumerable<String> fukas = from item in xml.Elements("コア情報").Elements("風化の程度区分判定表").Elements("風化の程度区分判定表_記号")
										select item.Value;

			IEnumerable<String> nessuis = from item in xml.Elements("コア情報").Elements("熱水変質の程度区分判定表").Elements("熱水変質の程度区分判定表_記号")
										select item.Value;

			IEnumerable<String> kounans = from item in xml.Elements("コア情報").Elements("硬軟区分判定表").Elements("硬軟区分判定表_記号")
										select item.Value;

			IEnumerable<String> chudos = from item in xml.Elements("コア情報").Elements("相対密度稠度").Elements("相対稠度_コード")
										select item.Value;

			IEnumerable<String> chudo2s = from item in xml.Elements("コア情報").Elements("相対密度稠度").Elements("相対稠度_状態")
										select item.Value;

			IEnumerable<String> saisyuritus = from item in xml.Elements("コア情報").Elements("コア採取率").Elements("コア採取率_採取率")
										select item.Value;

			IEnumerable<String> kizis = from item in xml.Elements("コア情報").Elements("観察記事").Elements("観察記事_記事")
										select item.Value;

			// 貼り付け位置
			int row = 3;
			int col = 1;

			//IEnumerableを配列に変換
			string[] hyoukou_array = hyoukous.ToArray();

			List<String> tmp = new List<String>();
			List<String> sindoList = new List<String>();

			// 深度を深度リストに格納  A
			foreach (String sindo in sindos)
			{
				sindoList.Add(sindo);
				//row++;
			}

			// 色調の深度を深度リストに格納  B
			for (int i = 0; i < sindoList.Count; i++)
			{
				foreach (String sindosiki in sindo2s)
				{
					if (sindoList[i].ToString() == sindosiki) // 同じ深度があれば次
						continue;

					if (sindoList.Count - 1 == i && Decimal.Parse(sindoList[i].ToString()) < Decimal.Parse(sindosiki))
						tmp.Add(sindosiki);

					else if (Decimal.Parse(sindoList[i].ToString()) < Decimal.Parse(sindosiki) && Decimal.Parse(sindosiki) < Decimal.Parse(sindoList[i + 1].ToString()))
						tmp.Add(sindosiki);
				}

				if (tmp.Count > 0)
				{
					sindoList.InsertRange(i + 1, tmp);
					tmp = new List<string>(); // tmp初期化
				}
			}

			// 色調の深度を深度リストに格納  C
			for (int j = 0; j < sindoList.Count; j++)
			{
				foreach (String sindofu in sindo3s)
				{
					if (sindoList[j].ToString() == sindofu) // 同じ深度があれば次
						continue;

					if (sindoList.Count - 1 == j && Decimal.Parse(sindoList[j].ToString()) < Decimal.Parse(sindofu))
						tmp.Add(sindofu);

					else if (Decimal.Parse(sindoList[j].ToString()) < Decimal.Parse(sindofu) && Decimal.Parse(sindofu) < Decimal.Parse(sindoList[j + 1].ToString()))
						tmp.Add(sindofu);
				}

				if (tmp.Count > 0)
				{
					sindoList.InsertRange(j + 1, tmp);
					tmp = new List<string>(); // tmp初期化
				}
			}

			// ボーリング名
			row = 3;  //初期表示位置
			foreach (String boring in borings)
			{
				xlSheet.Cells[row, col] = (boring);
				row++;
			}

			//孔口標高
			row = 3;  //初期表示位置
			foreach (String hyoukou in hyoukous)
			{
				xlSheet.Cells[row, col + 1] = (hyoukou);
				row++;
			}

			//総削孔長
			row = 3;  //初期表示位置
			foreach (String sousakukouchou in sousakukouchous)
			{
				xlSheet.Cells[row, col + 2] = (sousakukouchou);
				row++;
			}

			//標高
			row = 3;  //初期表示位置
			for (int h = 0, j = 0; h < sindoList.Count; h++)
			{
				xlSheet.Cells[row, col + 3] = Decimal.Parse(hyoukou_array[j]) - Decimal.Parse(sindoList[h]);
				row++;
			}

			//深度
			row = 3;  //初期表示位置
			for (int x = 0; x < sindoList.Count; x++)
			{
				xlSheet.Cells[row, col + 4] = (sindoList[x]);
				row++;
			}

			//上端深度～下端深度
			row = 3;
			for (int y = 0; y < sindoList.Count; y++)
			{
				if (row == 3)  //if (xlSheet.Cells[row, col + 6])
				{
					xlSheet.Cells[row, col + 5] = "0～" + sindoList[y];
					row++;
				}

				else if (row != 3)
				{
					xlSheet.Cells[row, col + 5] = sindoList[y - 1] + "～" + sindoList[y];
					//[y-1]は今より一つ前の要素の中身を表示 [y]は現在の繰り返し回数分の要素を表示
					row++;
				}
			}

			//層厚
			row = 3;
			for (int z = 0; z < sindoList.Count; z++)
			{
				if (row == 3)
				{
					xlSheet.Cells[row, col + 6] = sindoList[z];
					row++;
				}

				else if (row != 3)
				{
					xlSheet.Cells[row, col + 6] = Decimal.Parse(sindoList[z]) - Decimal.Parse(sindoList[z - 1]);
					row++;
				}
			}

			// 工学的地質区分名のデータ
			row = 3;  //初期表示位置
			foreach (String dositumei in dositumeis)
			{
				xlSheet.Cells[row, col + 7] = (dositumei);
				row++;
			}

			//色調のデータ
			row = 3;  //初期表示位置
			foreach (String sikichou in sikichous)
			{
				xlSheet.Cells[row, col + 8] = (sikichou);
				row++;
			}

			//風化の程度のデータ
			row = 3;  //初期表示位置
			foreach (String fuka in fukas)
			{
				xlSheet.Cells[row, col + 9] = (fuka);
				row++;
			}

			//熱水変質の程度のデータ
			row = 3;  //初期表示位置
			foreach (String nessui in nessuis)
			{
				xlSheet.Cells[row, col + 10] = (nessui);
				row++;
			}

			//破砕度  ----------------------------------

			//相対稠度のデータ
			row = 3;  //初期表示位置
			foreach (String chudo in chudos)
			{
				xlSheet.Cells[row, col + 12] = (chudo);
				row++;
			}

			//コア採取率のデータ
			row = 3;  //初期表示位置
			foreach (String saisyuritu in saisyuritus)
			{
				xlSheet.Cells[row, col + 13] = (saisyuritu);
				row++;
			}

			//記事のデータ
			row = 3;  //初期表示位置
			foreach (String kizi in kizis)
			{
				xlSheet.Cells[row, col + 14] = (kizi);
				row++;
			}

			//コア質量  ---------------------------------
		}


		/// <summary>
		/// セルダブルクリック時処理
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		// <param name="ws1"></param>
		private void btnZisuberiHyou_Click(object sender, EventArgs e)//, Excel.Worksheet ws1
		{
			#region "work 貼り付けるXMLを作成"

			#endregion

			// Excel操作用オブジェクト
			Microsoft.Office.Interop.Excel.Application xlApp = null;
			Microsoft.Office.Interop.Excel.Workbooks xlBooks = null;
			Microsoft.Office.Interop.Excel.Workbook xlBook = null;
			Microsoft.Office.Interop.Excel.Sheets xlSheets = null;
			Microsoft.Office.Interop.Excel.Worksheet xlSheet = null;

			// オープンファイルダイアログを生成する
			OpenFileDialog op = new OpenFileDialog();
			op.Title = "ファイルを開く";
			op.InitialDirectory = @"C:\";
			op.FileName = @"";
			op.Filter = "xmlファイル(*.xml;*)|*.xml;* | すべてのファイル(*.*)|*.*";
			op.FilterIndex = 1;

			//オープンファイルダイアログを表示する
			DialogResult result = op.ShowDialog();

			//拡張子判別
			string fileExt = System.IO.Path.GetExtension(op.FileName);

			if (result == DialogResult.OK)
			{
				//「開く」ボタンが選択された時の処理
				string fileName = op.FileName;  //ファイルパス取得
			}
			else if (result == DialogResult.Cancel)
			{
				//「キャンセル」ボタンまたは「×」ボタンが選択された時の処理
				return;
			}

			//待機状態のマウス・カーソルを表示
			Cursor preCursor = Cursor.Current;

			

			try  //例外が発生する可能性のある処理
			{
				String fName = "地すべりボーリング柱状図(標準貫入試験用)" + ".xlsx";

				string spf = System.Windows.Forms.Application.StartupPath + "\\ExcelTemplate"; // exeのパス取得

				if (!File.Exists(Path.Combine(spf, fName)))  //テンプレートファイルのチェック
				{
					MessageBox.Show("テンプレートファイルが存在しません。");
					return;
				}
				xlApp = new Microsoft.Office.Interop.Excel.Application(); // Excelアプリケーション生成
				xlApp.Visible = true;// 表示
				xlApp.DisplayAlerts = true; // エクセル保存確認表示
				// ◆新規のExcelブックを開く◆
				xlBook = xlApp.Workbooks.Add(spf + "\\" + fName);
				xlSheet = (Excel.Worksheet)xlBook.Worksheets[1];

				// マウスカーソルを待機状態にする
				Cursor.Current = Cursors.WaitCursor;
				// 待機ダイアログ表示
				WaitSplash.ShowSplash();

			}
			catch (Exception ie)  //例外が発生した時の処理
			{
				// xlSheet解放
				if (xlSheet != null)
				{
					Marshal.ReleaseComObject(xlSheet);
					xlSheet = null;
				}

				// xlSheets解放
				if (xlSheets != null)
				{
					Marshal.ReleaseComObject(xlSheets);
					xlSheets = null;
				}

				// xlBook解放
				if (xlBook != null)
				{
					try
					{
						xlBook.Close();
					}
					finally
					{
						Marshal.ReleaseComObject(xlBook);
						xlBook = null;
					}
				}

				// xlBooks解放
				if (xlBooks != null)
				{
					Marshal.ReleaseComObject(xlBooks);
					xlBooks = null;
				}

				// xlApp解放
				if (xlApp != null)
				{
					try
					{
						// アラートを戻して終了
						xlApp.DisplayAlerts = true;
						xlApp.Quit();
					}
					finally
					{
						Marshal.ReleaseComObject(xlApp);
					}
				}

				throw ie;
			}

			finally  //例外の有無にかかわらず、必ず最後に実行される処理
			{
				// 待機ダイアログ終了
				WaitSplash.CloseSplash();

				// マウスカーソルを元の状態に戻す
				Cursor.Current = preCursor;
			}

			//罫線
			Excel.Range range_line;
			range_line = xlSheet.Range[xlSheet.Cells[2, 1], xlSheet.Cells[60, 17]];

			Excel.Borders borders;  //型宣言
			Excel.Border border;  //型宣言
			//罫線の設定
			borders = range_line.Borders;
			border = borders[Excel.XlBordersIndex.xlEdgeLeft];
			border.LineStyle = Excel.XlLineStyle.xlContinuous;
			border = borders[Excel.XlBordersIndex.xlEdgeRight];
			border.LineStyle = Excel.XlLineStyle.xlContinuous;
			border = borders[Excel.XlBordersIndex.xlEdgeTop];
			border.LineStyle = Excel.XlLineStyle.xlContinuous;
			border = borders[Excel.XlBordersIndex.xlEdgeBottom];
			border.LineStyle = Excel.XlLineStyle.xlContinuous;

			//内枠
			border = borders[Excel.XlBordersIndex.xlInsideVertical];
			border.LineStyle = Excel.XlLineStyle.xlContinuous;
			border = borders[Excel.XlBordersIndex.xlInsideHorizontal];
			border.LineStyle = Excel.XlLineStyle.xlContinuous;

			// XMLファイルの読み込み
			XElement xml = XElement.Load(op.FileName);

			//それぞれのタグ内の情報を取得する
			IEnumerable<String> borings = from item in xml.Elements("標題情報").Elements("調査基本情報").Elements("ボーリング名")
										select item.Value;

			IEnumerable<String> hyoukous = from item in xml.Elements("標題情報").Elements("ボーリング基本情報").Elements("孔口標高")
										select item.Value;

			IEnumerable<String> sousakukouchous = from item in xml.Elements("標題情報").Elements("ボーリング基本情報").Elements("総削孔長")
										select item.Value;

			IEnumerable<String> sindos = from item in xml.Elements("コア情報").Elements("工学的地質区分名現場土質名").Elements("工学的地質区分名現場土質名_下端深度")
										select item.Value;

			IEnumerable<String> sindo2s = from item in xml.Elements("コア情報").Elements("色調").Elements("色調_下端深度")
										select item.Value;

			IEnumerable<String> sindo3s = from item in xml.Elements("コア情報").Elements("風化の程度区分").Elements("風化の程度区分_下端深度")
										select item.Value;

			IEnumerable<String> sindo4s = from item in xml.Elements("コア情報").Elements("熱水変質の程度区分").Elements("熱水変質の程度区分_下端深度")
										select item.Value;

			IEnumerable<String> sindo5s = from item in xml.Elements("コア情報").Elements("破砕度").Elements("破砕度_下端深度")
										select item.Value;

			IEnumerable<String> sindo6s = from item in xml.Elements("コア情報").Elements("硬軟区分").Elements("硬軟区分_下端深度")
										select item.Value;

			IEnumerable<String> sindo7s = from item in xml.Elements("コア情報").Elements("相対密度稠度").Elements("相対密度稠度_下端深度")
										select item.Value;

			IEnumerable<String> sindo8s = from item in xml.Elements("コア情報").Elements("観察記事").Elements("観察記事_下端深度")
										select item.Value;

			IEnumerable<String> sindo9s = from item in xml.Elements("コア情報").Elements("標準貫入試験").Elements("標準貫入試験_開始深度")
										select item.Value;

			IEnumerable<String> dositumeis = from item in xml.Elements("コア情報").Elements("工学的地質区分名現場土質名").Elements("工学的地質区分名現場土質名_工学的地質区分名現場土質名")
										select item.Value;

			IEnumerable<String> sikichous = from item in xml.Elements("コア情報").Elements("色調").Elements("色調_色調名")
										select item.Value;

			IEnumerable<String> fukas = from item in xml.Elements("コア情報").Elements("風化の程度区分判定表").Elements("風化の程度区分判定表_記号")
										select item.Value;

			IEnumerable<String> nessuis = from item in xml.Elements("コア情報").Elements("熱水変質の程度区分判定表").Elements("熱水変質の程度区分判定表_記号")
										select item.Value;

			IEnumerable<String> kounans = from item in xml.Elements("コア情報").Elements("硬軟区分判定表").Elements("硬軟区分判定表_記号")
										select item.Value;

			IEnumerable<String> chudos = from item in xml.Elements("コア情報").Elements("相対密度稠度").Elements("相対稠度_コード")
										select item.Value;

			IEnumerable<String> chudo2s = from item in xml.Elements("コア情報").Elements("相対密度稠度").Elements("相対稠度_状態")
										select item.Value;

			IEnumerable<String> mitudos = from item in xml.Elements("コア情報").Elements("相対密度稠度").Elements("相対密度_コード")
										select item.Value;

			IEnumerable<String> mitudo2s = from item in xml.Elements("コア情報").Elements("相対密度稠度").Elements("相対密度_状態")
										select item.Value;

			IEnumerable<String> kizis = from item in xml.Elements("コア情報").Elements("観察記事").Elements("観察記事_記事")
										select item.Value;

			IEnumerable<String> hyozyuns = from item in xml.Elements("コア情報").Elements("標準貫入試験").Elements("標準貫入試験_合計打撃回数")
										select item.Value;

			//IEnumerableを配列に変換
			string[] hyoukou_array = hyoukous.ToArray();

			//リスト作成
			List<String> tmp = new List<String>();
			List<String> sindoList = new List<String>();

			// 貼り付け位置
			int row = 3;
			int col = 1;

			// 深度を深度リストに格納  A
			foreach (String sindo in sindos)
			{
				sindoList.Add(sindo);
			}

			// 色調の深度を深度リストに格納  B
			for (int i = 0; i < sindoList.Count; i++)
			{
				foreach (String sindosiki in sindo2s)
				{
					if (sindoList[i].ToString() == sindosiki) // 同じ深度があれば次
						continue;

					if (sindoList.Count - 1 == i && Decimal.Parse(sindoList[i].ToString()) < Decimal.Parse(sindosiki))
						tmp.Add(sindosiki);

					else if (Decimal.Parse(sindoList[i].ToString()) < Decimal.Parse(sindosiki) && Decimal.Parse(sindosiki) < Decimal.Parse(sindoList[i + 1].ToString()))
						tmp.Add(sindosiki);
				}

				if (tmp.Count > 0)
				{
					sindoList.InsertRange(i + 1, tmp);
					tmp = new List<string>(); // tmp初期化
				}
			}

			// ボーリング名
			row = 3;  //初期表示位置
			foreach (String boring in borings)
			{
				xlSheet.Cells[row, col] = (boring);
				row++;
			}

			//孔口標高
			row = 3;  //初期表示位置
			foreach (String hyoukou in hyoukous)
			{
				xlSheet.Cells[row, col + 1] = (hyoukou);
				row++;
			}

			//総削孔長
			row = 3;  //初期表示位置
			foreach (String sousakukouchou in sousakukouchous)
			{
				xlSheet.Cells[row, col + 2] = (sousakukouchou);
				row++;
			}

			//標高
			row = 3;  //初期表示位置
			for (int h = 0, j = 0; h < sindoList.Count; h++)
			{
				xlSheet.Cells[row, col + 3] = Decimal.Parse(hyoukou_array[j]) - Decimal.Parse(sindoList[h]);
				row++;
			}

			//深度のデータ
			row = 3;  //初期表示位置
			for (int x = 0; x < sindoList.Count; x++)
			{
				xlSheet.Cells[row, col + 4] = (sindoList[x]);
				row++;
			}

			//上端深度～下端深度
			row = 3;
			for (int y = 0; y < sindoList.Count; y++)
			{
				if (row == 3)  //if (xlSheet.Cells[row, col + 6])
				{
					xlSheet.Cells[row, col + 5] = "0～" + sindoList[y];
					row++;
				}

				else if (row != 3)
				{
					xlSheet.Cells[row, col + 5] = sindoList[y - 1] + "～" + sindoList[y];
					//[y-1]は今より一つ前の要素の中身を表示 [y]は現在の繰り返し回数分の要素を表示
					row++;
				}
			}

			//層厚
			row = 3;
			for (int z = 0; z < sindoList.Count; z++)
			{
				if (row == 3)
				{
					xlSheet.Cells[row, col + 6] = sindoList[z];
					row++;
				}

				else if (row != 3)
				{
					xlSheet.Cells[row, col + 6] = Decimal.Parse(sindoList[z]) - Decimal.Parse(sindoList[z - 1]);
					row++;
				}
			}

			// 工学的地質区分名のデータ
			row = 3;  //初期表示位置
			foreach (String dositumei in dositumeis)
			{
				xlSheet.Cells[row, col + 7] = (dositumei);
				row++;
			}

			// 色調のデータ
			row = 3;  //初期表示位置
			foreach (String sikichou in sikichous)
			{
				xlSheet.Cells[row, col + 8] = (sikichou);
				row++;
			}

			// 風化の程度のデータ
			row = 3;  //初期表示位置
			foreach (String fuka in fukas)
			{
				xlSheet.Cells[row, col + 9] = (fuka);
				row++;
			}

			//熱水変質の程度のデータ
			row = 3;  //初期表示位置
			foreach (String nessui in nessuis)
			{
				xlSheet.Cells[row, col + 10] = (nessui);
				row++;
			}

			//硬軟のデータ
			row = 3;  //初期表示位置
			foreach (String kounan in kounans)
			{
				xlSheet.Cells[row, col + 11] = (kounan);
				row++;
			}

			//相対稠度のデータ
			row = 3;  //初期表示位置
			foreach (String chudo in chudos)
			{
				xlSheet.Cells[row, col + 12] = (chudo);
				row++;
			}

			//相対密度のデータ
			row = 3;  //初期表示位置
			foreach (String mitudo in mitudos)
			{
				xlSheet.Cells[row, col + 13] = (mitudo);
				row++;
			}

			// 記事のデータ
			row = 3;  //初期表示位置
			foreach (String kizi in kizis)
			{
				xlSheet.Cells[row, col + 14] = (kizi);
				row++;
			}

			//標準貫入試験(N値)のデータ
			row = 3;  //初期表示位置
			foreach (String hyozyun in hyozyuns)
			{
				xlSheet.Cells[row, col + 15] = (hyozyun);
				row++;
			}
		}

		private void btnClose_Click(object sender, EventArgs e)
		{
			this.Close();  //閉じる
		}
	}
}