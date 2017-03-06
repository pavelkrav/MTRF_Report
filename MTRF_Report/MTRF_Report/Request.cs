using System;
using System.Text;
using System.Net;
using System.Xml;
using System.Configuration;
using System.IO;

namespace MTRF_Report
{
	class Request
	{
		public string sdp_status { get; set; }  // "Failed" if no such request; otherwise "Success"

		public int workorderid { get; set; }
		public string requester { get; set; }
		public string createdby { get; set; }
		public long createdtime { get; set; }
		public long resolvedtime { get; set; }
		public string shortdescription { get; set; }
		public string timespentonreq { get; set; }
		public string subject { get; set; }
		public string site { get; set; }
		public string category { get; set; }
		public string subcategory { get; set; }
		public string technician { get; set; }
		public string status { get; set; }
		public string priority { get; set; }
		public string group { get; set; }
		public string description { get; set; }
		public string area { get; set; }

		public string resolution { get; set; }
		public long workMinutes { get; set; }

		public Request(int reqID)
		{
			WebClient wc = new WebClient();
			wc.Encoding = Encoding.UTF8;
			wc.Headers["Content-Type"] = "application/xml; charset=UTF-8";

			string reqStr = ConfigurationManager.AppSettings["SDP_PATH"] + "/request/" + reqID.ToString() + "?OPERATION_NAME=GET_REQUEST&TECHNICIAN_KEY=" + ConfigurationManager.AppSettings["SDP_API_KEY"];

			string xmlReqStr = null;
			try
			{
				xmlReqStr = wc.DownloadString(reqStr);
			}
			catch
			{
				xmlReqStr = "";
			}

			if (xmlReqStr.Length > 0)
			{

				try
				{
					using (XmlReader reader = XmlReader.Create(new StringReader(xmlReqStr)))
					{
						reader.ReadToFollowing("status");
						sdp_status = reader.ReadElementContentAsString();

						if (sdp_status == "Success")
						{
							reader.ReadToFollowing("value");
							workorderid = reader.ReadElementContentAsInt();
							reader.ReadToFollowing("value");
							requester = reader.ReadElementContentAsString();
							reader.ReadToFollowing("value");
							createdby = reader.ReadElementContentAsString();
							reader.ReadToFollowing("value");
							createdtime = reader.ReadElementContentAsLong();
							reader.ReadToFollowing("value");
							reader.ReadToFollowing("value");
							reader.ReadToFollowing("value");
							reader.ReadToFollowing("value");
							reader.ReadToFollowing("value");
							resolvedtime = reader.ReadElementContentAsLong();
							reader.ReadToFollowing("value");
							reader.ReadToFollowing("value");
							shortdescription = reader.ReadElementContentAsString();
							reader.ReadToFollowing("value");
							timespentonreq = reader.ReadElementContentAsString();
							reader.ReadToFollowing("value");
							subject = reader.ReadElementContentAsString();
							reader.ReadToFollowing("value");
							reader.ReadToFollowing("value");
							reader.ReadToFollowing("value");
							reader.ReadToFollowing("value");
							reader.ReadToFollowing("value");
							site = reader.ReadElementContentAsString();
							reader.ReadToFollowing("value");
							category = reader.ReadElementContentAsString();
							reader.ReadToFollowing("value");
							subcategory = reader.ReadElementContentAsString();
							reader.ReadToFollowing("value");
							technician = reader.ReadElementContentAsString();
							reader.ReadToFollowing("value");
							status = reader.ReadElementContentAsString();
							reader.ReadToFollowing("value");
							priority = reader.ReadElementContentAsString();
							reader.ReadToFollowing("value");
							reader.ReadToFollowing("value");
							reader.ReadToFollowing("value");
							group = reader.ReadElementContentAsString();
							reader.ReadToFollowing("value");
							description = reader.ReadElementContentAsString();
							reader.ReadToFollowing("value");
							area = reader.ReadElementContentAsString();

							if (site.Length < 3)
								site = null;
						}
						else
						{
							workorderid = reqID;
							status = null;
						}

						if (status == "Выполнено")
						{
							resolution = getResolution();
							workMinutes = getWorkTime();
						}
						else
						{
							resolution = null;
							workMinutes = 0;
						}
					}
				}
				catch (Exception e)
				{
					Console.WriteLine(e.Message);
				}
			}
		}

		private long getWorkTime()
		{
			WebClient wc = new WebClient();
			wc.Encoding = Encoding.UTF8;
			wc.Headers["Content-Type"] = "application/xml; charset=UTF-8";

			string reqStr = ConfigurationManager.AppSettings["SDP_PATH"] + "/request/" + workorderid.ToString() + "/worklogs?OPERATION_NAME=GET_WORKLOGS&TECHNICIAN_KEY=" + ConfigurationManager.AppSettings["SDP_API_KEY"];

			string xmlStr = null;
			try
			{
				xmlStr = wc.DownloadString(reqStr);
			}
			catch
			{
				xmlStr = "";
			}

			if (xmlStr.Length > 0)
			{

				try
				{
					using (XmlReader reader = XmlReader.Create(new StringReader(xmlStr)))
					{
						reader.ReadToFollowing("status");
						string status = reader.ReadElementContentAsString();

						if (status == "Success")
						{
							string param_name = null;
							do
							{
								reader.ReadToFollowing("name");
								param_name = reader.ReadElementContentAsString();
							}
							while (param_name != "workMinutes");
							reader.ReadToFollowing("value");
							return reader.ReadElementContentAsLong();
						}
						else return 0;
					}
				}
				catch
				{
					return 0;
				}
			}
			else return 0;
		}

		private string getResolution()
		{
			WebClient wc = new WebClient();
			wc.Encoding = Encoding.UTF8;
			wc.Headers["Content-Type"] = "application/xml; charset=UTF-8";

			string reqStr = ConfigurationManager.AppSettings["SDP_PATH"] + "/request/" + workorderid.ToString() + "?OPERATION_NAME=GET_RESOLUTION&TECHNICIAN_KEY=" + ConfigurationManager.AppSettings["SDP_API_KEY"];

			string xmlStr = null;
			try
			{
				xmlStr = wc.DownloadString(reqStr);
			}
			catch
			{
				xmlStr = "";
			}

			if (xmlStr.Length > 0)
			{

				try
				{
					using (XmlReader reader = XmlReader.Create(new StringReader(xmlStr)))
					{
						reader.ReadToFollowing("status");
						string status = reader.ReadElementContentAsString();

						if (status == "Success")
						{
							reader.ReadToFollowing("resolution");
							return reader.ReadElementContentAsString();
						}
						else return null;
					}
				}
				catch
				{
					return null;
				}
			}
			else return null;
		}

		public static DateTime longToDateTime(long dateNumber)
		{
			long beginTicks = new DateTime(1970, 1, 1, 3, 0, 0, DateTimeKind.Utc).Ticks;
			return new DateTime(beginTicks + dateNumber * 10000);
		}

		public string timeSpentRus()
		{
			string timeSpent = timespentonreq;
			timeSpent = timeSpent.Replace("hrs", "ч");
			timeSpent = timeSpent.Replace("min", "мин");
			return timeSpent;
		}

		public string technicianAcronym()
		{
			string[] array = technician.Split(' ');
			if (array.Length != 3)
				return technician;
			else
			{
				return $"{array[0]} {array[1].Substring(0, 1)}.{array[2].Substring(0, 1)}.";
			}
		}

		public string requesterAcronym()
		{
			string[] array = requester.Split(' ');
			if (array.Length != 3)
				return requester;
			else
			{
				return $"{array[0]} {array[1].Substring(0, 1)}.{array[2].Substring(0, 1)}.";
			}
		}

		public string readableSite()
		{
			string result = "";
			if (site != null)
				result += site;
			if (area == "Рождественка")
				return result;
			if (area == "РосГраница")
				result += "(РГ)";
			if (area == "ФАЖТ (Ст.Басманная)")
				result += "(ФАЖТ)";
			return result;
		}

		public static string convertFromHTML(string text)
		{
			if (String.IsNullOrWhiteSpace(text))
				return text;
			string temp = text;
			temp = temp.Replace("\n", ". ");
			// special symbols
			temp = temp.Replace("&quot;", "\"");
			temp = temp.Replace("&nbsp;", " ");
			temp = temp.Replace("&lt;", "<");
			temp = temp.Replace("&gt;", ">");
			// break row
			temp = temp.Replace("<div>", ". ");
			temp = temp.Replace("</div>", "");
			temp = temp.Replace("<p>", ". ");
			temp = temp.Replace("</p>", "");
			temp = temp.Replace("<br>", ". ");
			// other tags
			int openBr = 0;
			int closeBr = 0;
			string result = null;
			for (int i = 0; i < temp.Length; i++)
			{
				if (temp[i] == '<')
					openBr++;
				else if (temp[i] == '>')
					closeBr++;
				if (openBr == 0)
					result += temp[i];
				else if (openBr == closeBr)
				{
					openBr = 0;
					closeBr = 0;
				}
			}

			return result;
		}

		public void consoleOutput()
		{
			if (sdp_status == "Success")
			{
				Console.WriteLine($"Request ID: {workorderid}");
				Console.WriteLine($"Requester: {requester}");
				Console.WriteLine($"Created by: {createdby}");
				Console.WriteLine($"Subject: {subject}");
				Console.WriteLine($"Category: {category}");
				Console.WriteLine($"Subcategory: {subcategory}");
				Console.WriteLine($"Short description: {shortdescription}");
				Console.WriteLine($"Technician: {technician}");
				Console.WriteLine($"Created time: {longToDateTime(createdtime)}");
				if (resolvedtime > 0)
				{
					Console.WriteLine($"Resolved time: {longToDateTime(resolvedtime)}");
					Console.WriteLine($"Time spent: {timespentonreq}");
				}
				Console.WriteLine($"Status: {status}");
				Console.WriteLine($"Priority: {priority}");
				Console.WriteLine($"Group: {group}");
				Console.WriteLine($"Area: {area}");
				if (status == "Выполнено")
				{
					Console.WriteLine($"Resolution: {resolution}");
					Console.WriteLine($"Work time: {workMinutes/60000} min.");
					Console.WriteLine($"Worker in: {longToDateTime(resolvedtime - workMinutes).ToString(@"dd/MM/yyyy hh-mm")}");
				}
			}
			else
			{
				Console.WriteLine($"Request #{workorderid} does not exist.");
			}
		}

		public void consoleOutputMTRF()
		{
			if (sdp_status == "Success")
			{
				if (status == "Выполнено")
				{
					Console.WriteLine($"Date: {longToDateTime(resolvedtime).ToString(@"dd/MM/yyyy")}");
					Console.WriteLine($"Site: {readableSite()}");
					Console.WriteLine($"Requester: {requesterAcronym()}");
					Console.WriteLine($"Subject: {subject}");
					Console.WriteLine($"Resolution: {convertFromHTML(resolution)}");
					Console.WriteLine($"Technician: {technicianAcronym()}");
					Console.WriteLine($"Worker in: {longToDateTime(resolvedtime - workMinutes).ToString(@"HH:mm")}");
					Console.WriteLine($"Closed at: {longToDateTime(resolvedtime).ToString(@"HH:mm")}");
					Console.WriteLine($"Time spent: {timeSpentRus()}");
				}
				else
				{
					Console.WriteLine($"Request #{workorderid} is unresolved.");
				}
			}
			else
			{
				Console.WriteLine($"Request #{workorderid} does not exist.");
			}
		}

	}
}