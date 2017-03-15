using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Threading;
using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Command;
using XlsToApi.Model;
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
namespace XlsToApi.ViewModel
{
	public class MainViewModel : ViewModelBase
	{
		/// <summary>
		/// Initializes a new instance of the MainViewModel class.
		/// </summary>
		public MainViewModel(Dispatcher dispatcher)
		{
			ClickLoad = new RelayCommand(Load);
			ClickSend = new RelayCommand(Send);
			_excelApp = new Excel.Application();
			_dispatcher = dispatcher;
		}

		readonly Excel.Application _excelApp;
		private readonly Dispatcher _dispatcher;

		public RelayCommand ClickLoad { get; private set; }
		public RelayCommand ClickSend { get; private set; }
		private List<ScheduleModel> Schedules { get; } = new List<ScheduleModel>();
		public ObservableCollection<string> Logs { get; } = new ObservableCollection<string>();
		private void DispatherThreadRun(Action action)
		{
			_dispatcher?.BeginInvoke(action);
		}
		private void Send()
		{
			var sendMsg = JsonConvert.SerializeObject(Schedules);
			Logs.Add(sendMsg);
			string result = string.Empty;
			using (var client = new WebClient())
			{
				client.Headers[HttpRequestHeader.ContentType] = "application/json";
				var url = "http://fmsabwdapi.azurewebsites.net/api";
				try
				{
					result = client.UploadString(url, "POST", sendMsg);
				}
				catch (Exception ex)
				{
					Logs.Add(ex.Message);
				}
			}
			Logs.Add(result);
		}

		private void Load()
		{
			Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog
			{
				FileName = "Excel",
				DefaultExt = ".xls",
				Filter = "Excel documents (.xls)|*.xls;*.xlsx"
			};
			if (dlg.ShowDialog() == true)
			{
				CreateModel(dlg.FileName);
			}
		}

		private void CreateModel(string filename)
		{
			try
			{
				Schedules.Clear();
				Task.Run(() => ReadBook(_excelApp.Workbooks.Open(filename)));
			}
			catch (Exception ex)
			{
				Logs.Add(ex.Message);
			}
		}

		private void ReadBook(Excel.Workbook workbook)
		{
			foreach (Excel.Worksheet worksheet in workbook.Worksheets)
			{
				if (worksheet.Name == "Выписки" || worksheet.Name == "3114") continue;
				var schedul = new ScheduleModel { group_name = worksheet.Name };

				var cell = worksheet.Range["A1"];//range.Cells[1, 1];
				var dates = Regex.Matches(cell.Value.ToString(), "(\\d{2}\\.\\d{2}\\.\\d{2})");
				var dt = new DateTime();
				if (dates.Count > 0 && DateTime.TryParse(dates[0].Value, out dt))
					schedul.starts_at = dt;
				if (dates.Count > 1 && DateTime.TryParse(dates[1].Value, out dt))
					schedul.ends_at = dt;
				DispatherThreadRun(() => Logs.Add($"Группа: {schedul.group_name} расписание c {schedul.starts_at:d} по {schedul.ends_at:d} "));
				ReadScheet(worksheet, schedul);
				Schedules.Add(schedul);
			}
			//DispatherThreadRun(() => Logs.Add(JsonConvert.SerializeObject(Schedules)));
		}

		private void ReadScheet(Excel.Worksheet worksheet, ScheduleModel schedul)
		{
			Excel.Range range = worksheet.UsedRange;
			var items = new List<ScheduleItem>();
			var scheduls = new List<ScheduleDay>();
			var weekDay = string.Empty;
			var schedulweekDay = new ScheduleDay();
			try
			{
				for (var column = 1; column < range.Cells.Columns.Count; column++)
				{
					if (column < 5) continue;
					try
					{
						if (!string.IsNullOrEmpty(worksheet.Cells[4, column].Value?.ToString() ?? ""))
						{
							if (items.Any())
							{
								schedulweekDay.items = items.ToArray();
								scheduls.Add(schedulweekDay);
							}
							weekDay = worksheet.Cells[4, column].Value.Trim(' ');
							if(weekDay == "Преподаватель") return;
							schedulweekDay = new ScheduleDay { week_day = GetWeekDay(weekDay) };
							items = new List<ScheduleItem>();
							DispatherThreadRun(() => Logs.Add($"{weekDay}:"));
						}
						var subject = string.Empty;
						var teachers = new List<string>();
						for (var row = 1; row < range.Cells.Rows.Count; row++)
						{
							if (row < 6) continue;
							var b = worksheet.Cells[row, 2].Value?.ToString() ?? "";
							try
							{
								if (!string.IsNullOrEmpty(b) && Regex.Match(b, "^[а-яА-Я- \\.]+$").Success)
								{
									teachers = new List<string>();
									MatchCollection theacherMatches = Regex.Matches(worksheet.Cells[row, 41].Value ?? "", "([А-Я][а-яА-Я- ]+[А-Я\\.]*)");
									if (theacherMatches.Count > 0)
										teachers.AddRange(from Match teach in theacherMatches select teach.Value);
									subject = b;
								}
								if (string.IsNullOrEmpty(worksheet.Cells[row, column].Value?.ToString() ?? "")) continue;
								var cabs = GetRoomModel(worksheet, row, column);
								if (!cabs.Any()) continue;
								int lesson;
								if (!int.TryParse(worksheet.Cells[5, column].Value?.ToString() ?? "", out lesson)) continue;
								var item = new ScheduleItem
								{
									subject = subject,
									lesson = lesson,
									classrooms = cabs,
									week_is_odd = worksheet.Cells[row, 4].Value == "н",
									week_day = GetWeekDay(weekDay),
									teachers = teachers.ToArray()
								};
								DispatherThreadRun(() => Logs.Add(item.ToString()));
								items.Add(item);
							}
							catch (Exception ex)
							{
								DispatherThreadRun(() => Logs.Add(ex.Message));
							}
							row++;
						}
					}
					catch (Exception ex)
					{
						DispatherThreadRun(() => Logs.Add(ex.Message));
					}
				}
				schedul.schedule = scheduls.ToArray();
			}
			catch (Exception ex)
			{
				DispatherThreadRun(() => Logs.Add(ex.Message));
			}
			//var cell = worksheet.Range[$"E{row}"];
		}

		private RoomModel[] GetRoomModel(Excel.Worksheet worksheet, int row, int column)
		{
			var cabs = new List<RoomModel>();
			try
			{
				string value = worksheet.Cells[row, column].Value.ToString();
				var arr = value.Split(',');
				foreach (var el in arr)
				{
					if(string.IsNullOrEmpty(el)) continue;
					var cabMatch = Regex.Match(el, "(^\\d+[а-я]*)|([а-яА-Я-]*-\\d*)");
					if (!cabMatch.Success) continue;
					var cab = new RoomModel { room = cabMatch.Value };
					cabs.Add(cab);
				}
			}
			catch (Exception ex)
			{
				DispatherThreadRun(() => Logs.Add(ex.Message));
			}
			return cabs.ToArray();
		}

		private int GetWeekDay(string weekDay)
		{
			switch (weekDay.ToLower())
			{
				case "понедельник": return 1;
				case "вторник": return 2;
				case "среда": return 3;
				case "четверг": return 4;
				case "пятница": return 5;
				case "суббота": return 6;
				default: return 0;
			}
		}
	}
}