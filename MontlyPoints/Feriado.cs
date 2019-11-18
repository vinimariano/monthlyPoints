using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MontlyPoints
{
	public class Feriado
	{
		public DateTime? dateFormated { get; set; }
		public string name { get; set; }
		public string date { get; set; }
	}

	public class MDLFeriado
	{
		public long CodFeriado { get; set; }
		public string Descricao { get; set; }
		public DateTime DataFeriado { get; set; }
	}
}
