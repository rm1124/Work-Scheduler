using System;
using System.IO;
using System.Collections.Generic;
using System.Diagnostics;
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

namespace Work_Scheduler
{
	/// <summary>
	/// Page1.xaml の相互作用ロジック
	/// </summary>
	public partial class Page1 : Page
	{
		private static Page2 page2 = null;
		public object SelectedFile = null;
		public Page1()
		{
			InitializeComponent();
		}

		private void Button_Click(Object sender, RoutedEventArgs e)
		{
			switch ((sender as Button).Name)
			{
				case "Next":
					SelectedFile = FileList.SelectedItem;
					if (SelectedFile != null)
					{
						if (page2 == null)
						{
							page2 = new Page2(SelectedFile);
						}
						Debug.WriteLine(SelectedFile + "が選択されました");
						this.NavigationService.Navigate(page2 , SelectedFile);
					}
					else
					{
						Console.Inlines.Clear();
						Console.Inlines.Add(new Run("err:ファイルが選択されていません"));
						Console.Foreground = Brushes.Red;
					}
					break;

				case "Get":
					try
					{
						FileList.Items.Clear();
						string[] names = Directory.GetFiles(@"C:\WPFapps\Work Scheduler\Work Scheduler\files", "*.xlsx");
						foreach (string name in names)
						{
							Debug.WriteLine(name);
							FileList.Items.Add(name);
						}
					}
					catch (Exception c)
					{
						Debug.WriteLine(c.ToString());
					}
					break;

			}
		}
	}
}




