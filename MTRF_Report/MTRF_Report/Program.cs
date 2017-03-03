using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MTRF_Report
{
	class Program
	{
		static void Main(string[] args)
		{
			Request req = new Request(8336);
			req.consoleOutputMTRF();
			Console.ReadKey();
		}
	}
}
