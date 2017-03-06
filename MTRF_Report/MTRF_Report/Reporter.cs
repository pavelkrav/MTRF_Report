using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Xml;
using System.Configuration;
using Microsoft.Office.Interop.Excel;

namespace MTRF_Report
{
	class Reporter
	{
		public bool creatingError { get; }

		public int reqAmount { get; protected set; }
		public int techAmount { get; }
		public string[] technicians { get; }

		public Reporter()
		{
			try
			{
				using (XmlReader reader = XmlReader.Create(@"ini.xml"))
				{
					reader.ReadToFollowing("ini");
					reader.ReadToFollowing("lastrequest");
					reqAmount = reader.ReadElementContentAsInt();

					reader.ReadToFollowing("technicians");
					reader.ReadToFollowing("amount");
					techAmount = reader.ReadElementContentAsInt();
					technicians = new string[techAmount];
					for (int i = 0; i < techAmount; i++)
					{
						reader.ReadToFollowing("technician");
						technicians[i] = reader.ReadElementContentAsString();
					}
				}
				creatingError = false;
				updateReqAmount();
			}
			catch (Exception)
			{
				Console.WriteLine("Could not find \"ini.xml\" file or it is not valid.");
				creatingError = true;
			}
		}

		public void updateReqAmount()
		{
			int newReqAmount = reqAmount;
			int err = 0;
			int i = 0;
			Request req = new Request(reqAmount);
			do
			{
				i++;
				Console.Write($"Request #{reqAmount + i} ");
				req = new Request(reqAmount + i);
				if (req.sdp_status == "Failed")
				{
					Console.Write("Failed\n");
					err++;
				}
				else if (req.sdp_status == "Success")
				{
					Console.Write("Success\n");
					newReqAmount = req.workorderid;
					err = 0;
				}
				else
					err = 100;
			}
			while (err < 30);
			Console.Clear();

			try
			{
				XmlDocument doc = new XmlDocument();
				doc.Load(@"ini.xml");
				XmlElement ini = doc.DocumentElement;
				XmlNode lastreq = ini.FirstChild;
				lastreq.RemoveChild(lastreq.FirstChild);
				lastreq.AppendChild(doc.CreateTextNode(newReqAmount.ToString()));

				doc.Save(@"ini.xml");
				Console.WriteLine("\"ini.xml\" file has been changed");
			}
			catch (Exception)
			{
				Console.WriteLine("Could not find \"ini.xml\" file or it is not valid.");
				Console.WriteLine("ini.xml file has not been changed");
			}
			reqAmount = newReqAmount;
			Console.WriteLine($"Last request is #{reqAmount}");
		}

		public void createTsvReport()
		{
			long week = 86400 * 7 * 1000;

			Request req = new Request(reqAmount);
			long lTime = req.createdtime;

			int i = reqAmount;

			List<Request> resolvedList = new List<Request>();
			resolvedList = new List<Request>();

			do
			{
				req = new Request(i);
				if (req.sdp_status == "Failed")
				{
					Console.WriteLine($"Checked request #{i} - Does not exist");
					i--;
				}
				else if (req.status != "Выполнено")
				{
					Console.WriteLine($"Checked request #{i} - Pending");
					i--;
				}
				else if (req.resolvedtime < lTime - week - 32400000) // 32400000 = 9 hours
				{
					Console.WriteLine($"Checked request #{i} - Resolved");
					i--;
				}
				else if (req.sdp_status == "Success" && req.resolvedtime > lTime - week - 32400000)
				{
					for (int j = 0; j < techAmount; j++)
					{
						if (req.technician == technicians[j])
						{
							resolvedList.Add(req);
						}
					}
					Console.WriteLine($"Checked request #{i} - Recently resolved");
					i--;
				}
			}
			while (req.createdtime > lTime - week * 5 || req.sdp_status == "Failed");   // checking for last 5 weeks

			string name = DateTime.Now.ToString(@"dd/MM/yyyy hh-mm") + ".tsv";
			string path = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\AppData\Local\Temp\SDP\Reports\";

			Directory.CreateDirectory(path);

			// Resolved requests
			using (StreamWriter sw = new StreamWriter(File.Open(path + name, FileMode.Create), Encoding.UTF32))
			{
				sw.WriteLine("№\tДата\tКабинет\tЗаявитель\tОписание заявки\tРешение\tПринял\tВремя прибытия инженера\tВремя закрытия заявки\tОбщее время выполнения");
				int num = 0;

				foreach (Request rq in resolvedList)
				{
					num++;
					sw.Write(num + "\t");
					sw.Write(Request.longToDateTime(rq.resolvedtime).ToString(@"dd/MM/yyyy") + "\t");
					sw.Write($"{rq.readableSite()}\t");
					sw.Write(rq.requesterAcronym() + "\t");
					sw.Write(Request.convertFromHTML(rq.subject) + "\t");
					sw.Write(Request.convertFromHTML(rq.resolution) + "\t");
					sw.Write(rq.technicianAcronym() + "\t");
					sw.Write($"{Request.longToDateTime(rq.resolvedtime - rq.workMinutes).ToString(@"HH:mm")}\t");
					sw.Write($"{Request.longToDateTime(rq.resolvedtime).ToString(@"HH:mm")}\t");
					sw.Write($"{rq.timeSpentRus()}\n");
				}
				sw.Close();
			}

			Console.Clear();
			Console.WriteLine("Generated report file: " + path + name);

			Application excel = new Application();

			Workbook wb = excel.Workbooks.Open(path + name);
			Worksheet ws1 = (Worksheet)wb.Worksheets[1];
			ws1.Name = "Отчет";
			ws1.Columns.AutoFit();
			Range rng = ws1.Range["A1", "A2"].EntireColumn;
			rng.Font.Bold = true;
			rng = ws1.Range["A1", "B1"].EntireRow;
			rng.Font.Bold = true;

			excel.Visible = true;
		}

	}
}
