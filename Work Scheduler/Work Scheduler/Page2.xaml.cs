using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace Work_Scheduler
{
	/// <summary>
	/// Page2.xaml の相互作用ロジック
	/// </summary>
	public partial class Page2 : Page
	{
		public Excel.Application ExcelApp;
		public Excel.Workbook WorkbookBk;
		public Page2(object FileName)
		{
			InitializeComponent();
			Console2.Inlines.Clear();
			Console2.Inlines.Add(new Run(FileName + "を編集中です\n終了する際は必ず下部の[終了]ボタンから行ってください"　));
			this.ExcelApp = new Excel.Application();
			this.ExcelApp.Visible = false;
			this.WorkbookBk = (Excel.Workbook)(ExcelApp.Workbooks.Open ((string)FileName,
				Type.Missing,
				Type.Missing,
				Type.Missing,
				Type.Missing,
				Type.Missing,			
				Type.Missing,
				Type.Missing,
				Type.Missing,
				Type.Missing,
				Type.Missing,
				Type.Missing,
				Type.Missing,
				Type.Missing,
				Type.Missing));
				

		}

		private void Button_Clickp2(Object sender, RoutedEventArgs e)
		{
			switch ((sender as Button).Name)
			{
				case "copy": //開始時刻の値を終了時刻欄にコピー
					string StartY = this.Year.Text;
					string StartMo = this.Month.Text;
					string StartD = this.Day.Text;
					string StartH = this.Hour.Text;
					string StartMi = this.Minute.Text;
					Debug.WriteLine(StartY + "年" + StartMo + "月" + StartD + "日" + StartH + "時" + StartMi + "分をコピー");

					this.EYear.Text = StartY;
					this.EMonth.Text = StartMo;
					this.EDay.Text = StartD;
					this.EHour.Text = StartH;
					this.EMinute.Text = StartMi;

					break;

				case "NowTime": //開始時刻欄に現在時刻を反映
					var dt = DateTime.Now;
					var arr = dt.ToString("yyyy年MM月dd日HH時mm分"); //取得したDateTimeをstring型に変換
					var dtList = Regex.Split(arr, "年|月|日|時|分"); //変換した値を分割して配列に収納する
					this.Year.Text = dtList[0];
					this.Month.Text = dtList[1];
					this.Day.Text = dtList[2];
					this.Hour.Text = dtList[3];
					this.Minute.Text = dtList[4];

					Debug.WriteLine("現在時刻:" + arr + "を反映しました。");
					break;

				case "Registration": //入力された値をセルに記入する

					if ((Year.SelectedItem != null || string.IsNullOrEmpty(Year.Text) == false) && (EYear.SelectedItem != null || string.IsNullOrEmpty(EYear.Text) == false) && (Month.SelectedItem != null || string.IsNullOrEmpty(Month.Text) == false) && (EMonth.SelectedItem != null || string.IsNullOrEmpty(EMonth.Text) == false) && (Day.SelectedItem != null || string.IsNullOrEmpty(Day.Text) == false) && (EDay.SelectedItem != null || string.IsNullOrEmpty(EDay.Text) == false) && (Hour.SelectedItem != null || string.IsNullOrEmpty(Hour.Text) == false) && (EHour.SelectedItem != null || string.IsNullOrEmpty(EHour.Text) == false) && (Minute.SelectedItem != null || string.IsNullOrEmpty(Minute.Text) == false) && (EMinute.SelectedItem != null || string.IsNullOrEmpty(EMinute.Text) == false)) //空欄がないか調べ、存在した場合後続処理をパスする
					{
						string RegY = this.Year.Text;
						string RegMo = this.Month.Text;
						string RegD = this.Day.Text;
						string RegH = this.Hour.Text;
						string RegMi = this.Minute.Text;
						string RegEY = this.EYear.Text;
						string RegEMo = this.EMonth.Text;
						string RegED = this.EDay.Text;
						string RegEH = this.EHour.Text;
						string RegEMi = this.EMinute.Text;

						if (DatingChecker(RegY, RegMo, RegD) == true && DatingChecker(RegEY, RegEMo, RegED) == true) //日付が存在するかチェックし、存在しなければ後続処理を行わない
						{
							string StartTime = (RegY + "/" + RegMo + "/" + RegD + " " + RegH + ":" + RegMi);
							DateTime STime = DateTime.Parse(StartTime);

							string EndTime = (RegEY + "/" + RegEMo + "/" + RegED + " " + RegEH + ":" + RegEMi);
							DateTime ETime = DateTime.Parse(EndTime);

							int WorkTime = TimeComparer(STime, ETime);
							this.EditLog.Items.Insert(0, "勤務時刻:" + RegY + "年" + RegMo + "/" + RegD + " " + RegH + ":" + RegMi + " ‐ " + RegEY + "年" + RegEMo + "/" + RegED + " " + RegEH + ":" + RegEMi + "を登録しました(勤務時間" + WorkTime / 3600 + "時間)");
							Excel.Worksheet EditSheet = (Excel.Worksheet)this.WorkbookBk.Sheets[1];
							string defaultRange = vsSearch("Day1");
							EditSheet.Range[RangeConverter(defaultRange, RegD, 0)].Value = RegH + ":" + RegMi;

							if (RegD != RegED) //勤務開始日と終了日が異なる場合の処理
							{
								int HourWorkTime = WorkTime / 3600;
								int RegNewEH = HourWorkTime + int.Parse(RegH);
								RegEH = "'" + RegNewEH.ToString();
								RegED = RegD;
							}
							EditSheet.Range[RangeConverter(defaultRange, RegED, 1)].Value = RegEH + ":" + RegEMi;
							EditSheet.Range[RangeConverter(defaultRange, RegD , 2)].Value = statusSearch("Break:");
							this.WorkbookBk.Save();
						}
						else
						{
							this.EditLog.Items.Insert(0, "日付の値が不正です");
						}
					}
					else
					{
						this.EditLog.Items.Insert(0, "空欄があります");
					}


					break;

				case "End": //アプリケーションの終了(Excelアプリケーションを終了する)
					this.WorkbookBk.Close();
					Application.Current.Shutdown();
					break;

				case "Load": //files配下のstatus.txtをシートに反映する
					Excel.Worksheet EditSheetSt = (Excel.Worksheet)this.WorkbookBk.Sheets[1];
					string RangeMo = vsSearch("Month:");
					string RangeNa = vsSearch("Name:");
					string RangeCont = vsSearch("Contract:");
					string RangeWo = vsSearch("WorkPlace:");
					string RangeSt = vsSearch("Start:");
					string RangeEn = vsSearch("End:");

					EditSheetSt.Range[RangeMo].Value = statusSearch("Month:");
					EditSheetSt.Range[RangeNa].Value = statusSearch("Name:");
					EditSheetSt.Range[RangeCont].Value = statusSearch("Contract:");
					EditSheetSt.Range[RangeWo].Value = statusSearch("WorkPlace:");
					EditSheetSt.Range[RangeSt].Value = statusSearch("Start:");
					EditSheetSt.Range[RangeEn].Value = statusSearch("End:");
					this.WorkbookBk.Save();
					this.EditLog.Items.Insert(0, "statusファイルをロードしました");
					break;
			}	
		}

		public int TimeComparer(DateTime STime , DateTime ETime) //UnixTimeを用いて開始時刻と終了時刻の差を求める
		{
			int diffTime = GetUnixTime(ETime) - GetUnixTime(STime);
			return diffTime;
		}

		public static int GetUnixTime(DateTime timeStamp) //UnixTimeの取得処理
		{
			var unixTimestamp = (int)(timeStamp.Subtract(new DateTime(1970, 1, 1))).TotalSeconds;
			return unixTimestamp;
		}

		public static string vsSearch(string SearchValue) //vs.txtよりキーを検索し返り値として対応する文字列を与える
		{
			StreamReader sr = new StreamReader(@"C:\WPFapps\Work Scheduler\Work Scheduler\files\vs.txt");
			string line = "";
			string GetValue = null;
			Debug.WriteLine(SearchValue);
			while ((line = sr.ReadLine())!= null)
			{
				int num = line.IndexOf(SearchValue);
				if (num >= 0)
				{
					Debug.WriteLine("sucsess");
					string sv = (string)line;
					GetValue = sv.Replace(SearchValue, "");
				}
			}
			return GetValue;
		}

		public static string statusSearch(string SearchValueSt) //status.txtよりキーを検索し返り値として対応する文字列を与える
		{
			StreamReader srSt = new StreamReader(@"C:\WPFapps\Work Scheduler\Work Scheduler\files\status.txt");
			string lineSt = "";
			string GetValueSt = null;
			Debug.WriteLine(SearchValueSt);
			while ((lineSt = srSt.ReadLine()) != null)
			{
				int numSt = lineSt.IndexOf(SearchValueSt);
				if (numSt >= 0)
				{
					Debug.WriteLine("sucsess");
					string svSt = (string)lineSt;
					GetValueSt = svSt.Replace(SearchValueSt, "");
				}
			}
			return GetValueSt;
		}

		public static string RangeConverter(string CellRange , string Time ,　int StartFlag) //vs.txtの値と日付からセル番地を計算
		{

			var NumAlpha = new Regex("(?<Alpha>[a-zA-Z]*)(?<Numelic>[0-9]+)"); //セル番地の値を分割
			var match = NumAlpha.Match(CellRange);
			int writeTime = int.Parse(match.Groups["Numelic"].Value); //数値を抜き出し
			Debug.WriteLine(writeTime);
			int newTime = writeTime + int.Parse(Time) - 1; //数値に日付を足して書き込むセルを指定
			string alphaNo = match.Groups["Alpha"].ToString();
			int index = 0;
			string alphabet = null;
			if (StartFlag == 1 || StartFlag == 2) //終業時刻の場合に列を1つずらす
			{
				index += StartFlag;
				for (int i = 0; i < alphaNo.Length; i++)
				{
					int num = Convert.ToChar(alphaNo[alphaNo.Length - i - 1]) - 64;

					index += (int)(num * Math.Pow(26, i));
				}

				index--;
				do
				{
					alphabet = Convert.ToChar(index % 26 + 0x41) + alphabet;
				}
				while
				((index = index / 26 - 1) != -1);
			}
			else
			{
				alphabet = alphaNo;
			}

			string WriteRange = alphabet + newTime.ToString();
			Debug.WriteLine(WriteRange);
			return WriteRange;
		}

		public static bool DatingChecker(string NewYear , string NewMonth , string NewDay) //日付の整合性をチェック
		{
			if((NewMonth == "04" || NewMonth == "06" || NewMonth == "09" || NewMonth == "11") && int.Parse(NewDay) <= 30)
			{
				return true;
			}
			else if(NewMonth == "02" && int.Parse(NewDay) <= 28)
			{
				return true;
			}
			else if(NewMonth == "02" && int.Parse(NewYear) % 4 == 0 && !(int.Parse(NewYear) % 100 == 0 && int.Parse(NewYear) % 400 != 0))
			{
				return true;
			}
			else if(int.Parse(NewDay) <= 31 )
			{
				return false;
			}
			else
			{
				return false;
			}
		}
	}
}
