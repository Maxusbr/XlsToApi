using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
namespace XlsToApi.Model
{
	public class ScheduleItem
	{
		public DateTime? date { get; set; }
		public TimeSpan lesson_start { get; set; }
		public TimeSpan lesson_end { get; set; }
		[JsonIgnore]
		public int lesson { get; set; }
		public string lesson_type { get; set; }
		public string subject { get; set; }
		public string[] teachers { get; set; }
		public RoomModel[] classrooms { get; set; }
		public bool week_is_odd { get; set; }
		public int week_day { get; set; }

		public override string ToString()
		{
			return $"{subject}: кабинеты({string.Join(",", classrooms.Select(o => o.room).ToArray())}), пара №{lesson}, " +
			       $"день({week_day}), нечетная({week_is_odd}), преподаватели({string.Join(",", teachers)})";
		}
	}
}
