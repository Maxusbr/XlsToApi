using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XlsToApi.Model
{
	public class ScheduleModel
	{
		public int id { get; set; }
		public string group_name { get; set; }
		public string faculty_name { get; set; }
		public DateTime starts_at { get; set; }
		public DateTime ends_at { get; set; }
		public bool is_session { get; set; }
		public ScheduleDay[] schedule { get; set; }
	}
}
