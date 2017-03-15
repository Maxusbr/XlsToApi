using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XlsToApi.Model
{
	public class ScheduleDay
	{
		public DateTime date { get; set; }
		public int week_day { get; set; }
		public ScheduleItem[] items { get; set; }
	}
}
