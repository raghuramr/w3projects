using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;

namespace TaxClaimers.WinAppGens
{
	public class Utilities
	{
		private static readonly Random rnd = new Random();
		private static readonly object syncLock = new object();

		public static string GetRandomNumber(int length)
		{
			string s = string.Empty;

			lock (syncLock)
			{
				s = rnd.Next(1, 9).ToString();

				for (int i = 0; i < length - 1; i++)
				{
					s = String.Concat(s, rnd.Next(10).ToString());
				}

				return s;
			}
		}

		/// <summary>
		/// Function to generate the random number for a given value by making Min and Max of contingency
		/// </summary>
		/// <param name="value"></param>
		/// <param name="contingency"></param>
		/// <returns></returns>
		public static string GetRandomNumberUsingContingency(int value, int contingency)
		{
			lock (syncLock)
			{
				return rnd.Next(value - contingency, value + contingency).ToString();
			}
		}

		public static string GetRandomTime()
		{
			lock (syncLock)
			{
				string time = string.Format("{0}:{1}:{2}", rnd.Next(1, 12), rnd.Next(0, 59), rnd.Next(0, 59));
				var result = Convert.ToDateTime(time);
				string takeThis = result.ToString("hh:mm:ss tt", CultureInfo.CurrentCulture);

				return takeThis;
			}
		}

		public static int GetRandomNumberBetween(int min, int max)
		{
			lock (syncLock)
			{
				return rnd.Next(min, max);
			}
		}

		public string NumberToEnglish(long n)
		{
			StringBuilder builder = new StringBuilder();
			var thousand = 1000;
			var hundred = 100;

			while (n != 0)
			{
				if (n >= thousand)
				{
					builder.AppendFormat("{0} Thousand ", NumberToEnglish(n / thousand));
					n -= (n / thousand) * thousand;
				}
				else if (n >= hundred)
				{
					builder.AppendFormat("{0} Hundred ", NumberToEnglish(n / hundred));
					n -= (n / hundred) * hundred;
				}
				else if (n >= 20)
				{
					builder.AppendFormat(numerals[(n / 10) * 10] + " ");
					n -= (n / 10) * 10;
				}
				else
				{
					builder.Append(numerals[n]);
					n -= n;
				}
			}

			return builder.ToString().ToLower();
		}

		Dictionary<long, string> numerals = new Dictionary<long, string>(){
			{1,"One"},
			{2,"Two"},
			{3,"Three"},
			{4,"Four"},
			{5,"Five"},
			{6,"Six"},
			{7,"Seven"},
			{8,"Eight"},
			{9,"Nine"},
			{10,"Ten"},
			{11,"Eleven"},
			{12,"Twelve"},
			{13,"Thirteen"},
			{14,"Fourteen"},
			{15,"Fifteen"},
			{16,"Sixteen"},
			{17,"Seventeen"},
			{18,"Eighteen"},
			{19,"Nineteen"},
			{20,"Twenty"},
			{30,"Thirty"},
			{40,"Forty"},
			{50,"Fifty"},
			{60,"Sixty"},
			{70,"Seventy"},
			{80,"Eighty"},
			{90,"Ninety"},
			{100,"Hundred"},
			{1000,"Thousand"}
		};

		string[] ones = { "", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten", "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen", "seventeen", "eighteen", "nineteen" };
		string[] tens = { "", "ten", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety" };
		string[] thou = { "", "thousand", "million", "billion", "trillion", "quadrillion", "quintillion" };

		public string NumberToWords(string rawnumber)
		{
			int inputNum = 0;
			int dig1, dig2, dig3, level = 0, lasttwo, threeDigits;
			string rupees, paisa;
			try
			{
				string[] Splits = new string[2];
				Splits = rawnumber.Split('.');   //notice that it is ' and not "
				inputNum = Convert.ToInt32(Splits[0]);
				rupees = "";
				paisa = Splits[1];
				if (paisa.Length == 1)
				{
					paisa += "0";   // 12.5 is twelve and 50/100, not twelve and 5/100
				}
			}
			catch
			{
				paisa = "00";
				inputNum = Convert.ToInt32(rawnumber);
				rupees = "";
			}

			string x = "";

			bool isNegative = false;
			if (inputNum < 0)
			{
				isNegative = true;
				inputNum *= -1;
			}
			if (inputNum == 0)
			{
				return "zero paisa";
			}

			string s = inputNum.ToString();

			while (s.Length > 0)
			{
				//Get the three rightmost characters
				x = (s.Length < 3) ? s : s.Substring(s.Length - 3, 3);

				// Separate the three digits
				threeDigits = int.Parse(x);
				lasttwo = threeDigits % 100;
				dig1 = threeDigits / 100;
				dig2 = lasttwo / 10;
				dig3 = (threeDigits % 10);

				// append a "thousand" where appropriate
				if (level > 0 && dig1 + dig2 + dig3 > 0)
				{
					rupees = thou[level] + " " + rupees;
					rupees = rupees.Trim();
				}

				// check that the last two digits is not a zero
				if (lasttwo > 0)
				{
					if (lasttwo < 20)
					{
						// if less than 20, use "ones" only
						rupees = ones[lasttwo] + " " + rupees;
					}
					else
					{
						// otherwise, use both "tens" and "ones" array
						rupees = tens[dig2] + " " + ones[dig3] + " " + rupees;
					}
					if (s.Length < 3)
					{
						if (isNegative) { rupees = "negative " + rupees; }
						return rupees + " rupees and " + getPaisaSpelled(paisa);
					}
				}

				// if a hundreds part is there, translate it
				if (dig1 > 0)
				{
					rupees = ones[dig1] + " hundred " + rupees;
					s = (s.Length - 3) > 0 ? s.Substring(0, s.Length - 3) : "";
					level++;
				}
				else
				{
					if (s.Length > 3)
					{
						s = s.Substring(0, s.Length - 3);
						level++;
					}
				}
			}
			
			if (isNegative) { rupees = "negative " + rupees; }
			return rupees + " rupees and" + getPaisaSpelled(paisa);
		}

		private string getPaisaSpelled(string _paisa)
		{
			string pisaSpelled ="";

			while (_paisa.Length > 0)
			{
				//Get the three rightmost characters
				string x = (_paisa.Length < 3) ? _paisa : _paisa.Substring(_paisa.Length - 3, 3);

				// Separate the three digits
				int threeDigits = int.Parse(x);
				int lasttwo = threeDigits % 100;
				int dig1 = threeDigits / 100;
				int dig2 = lasttwo / 10;
				int dig3 = (threeDigits % 10);

				// check that the last two digits is not a zero
				if (lasttwo > 0)
				{
					if (lasttwo < 20)
					{
						// if less than 20, use "ones" only
						pisaSpelled = ones[lasttwo] + " " + pisaSpelled;
					}
					else
					{
						// otherwise, use both "tens" and "ones" array
						pisaSpelled = tens[dig2] + " " + ones[dig3] + " " + pisaSpelled;
					}
					if (_paisa.Length < 3)
					{
						return pisaSpelled + "paisa only";
					}
				}
			}

			return pisaSpelled;
		}
	}
}
