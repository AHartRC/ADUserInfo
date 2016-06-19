using System;
using System.Collections.Generic;
using System.Configuration;
using System.DirectoryServices.AccountManagement;
using System.IO;
using System.Linq;

namespace ADUserInfo
{
	class Program
	{
		// C:\Users\anhart\Documents\FM_Workspace_User_List.txt
		private static string _filePath;
		private static readonly string OutputPath = string.Format(@"{0}\{1}", Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), string.Format("ADUserInfo_{0}.txt", DateTime.Now.ToFileTime()));

		private static long _inputCount;
		private static long OutputCount
		{
			get { return Users.Count; }
		}

		private static readonly ContextType ContextType = DefaultContextType;
		private static readonly string DomainName = ConfigurationManager.AppSettings["DomainName"] ?? "bcbg";
		private static readonly string DomainSuffix = ConfigurationManager.AppSettings["DomainSuffix"] ?? "com";
		private static readonly string Domain = string.Format("{0}.{1}", DomainName, DomainSuffix);
		private static readonly string DomainContainer = string.Format("DC={0},DC={1}", DomainName, DomainSuffix);

		private static readonly List<UserPrincipal> Users = new List<UserPrincipal>();

		static void Main(string[] args)
		{
			Console.WriteLine("********************************************************");
			Console.WriteLine("**     Active Directory User Information Retriever    **");
			Console.WriteLine("********************************************************");
			Console.WriteLine("Default: C:\\Users\\anhart\\Documents\\FM_Workspace_User_List.txt");
			Console.WriteLine("This application will take a text file that contains a\r\nlist of names (one per line) and will retrieve the\r\nuser information from active directory that matches\r\nthe users' names in the list.");
			Console.WriteLine("********************************************************");
			Console.WriteLine("Arguments:");
			foreach (string arg in args)
			{
				Console.WriteLine(arg);
			}
			Console.WriteLine("********************************************************");
			GetFilePath();
			ProcessFile();
			OutputResults();
			Console.ReadKey(true);
		}

		public static ContextType DefaultContextType
		{
			get
			{
				string contextTypeString = ConfigurationManager.AppSettings["ContextType"] ?? "domain";
				switch (contextTypeString.Trim().ToLower())
				{
					case "domain":
						return ContextType.Domain;
					case "applicationdirectory":
						return ContextType.ApplicationDirectory;
					default:
						return ContextType.Machine;
				}
			}
		}

		private static PrincipalContext GetPrincipalContext()
		{
			switch (ContextType)
			{
				case ContextType.Domain:
					return new PrincipalContext(ContextType, Domain, DomainContainer);
				default:
					return new PrincipalContext(ContextType);
			}
		}

		private static void GetFilePath()
		{
			Console.WriteLine("Please input the full path of the file you wish to parse:");
			// Read the path of the file to parse
			_filePath = Console.ReadLine();

			// If the path entered is null, empty or contains only whitespace, re-acquire the filePath
			if (string.IsNullOrWhiteSpace(_filePath) || !File.Exists(_filePath))
			{
				// Notify the user
				Console.WriteLine("The file path you input is invalid!");
				// Re-acquire file path
				GetFilePath();
			}
		}

		private static void ProcessFile()
		{
			PrincipalContext context = GetPrincipalContext();

			// Grab the file contents
			string[] lines = File.ReadAllLines(_filePath);
			_inputCount = lines.Count();

			foreach (string line in lines)
			{
				// Get the users' first and last name
				string[] names = line.Split(' ');

				// Instantiate the principal to search upon. This is not the actual user record.
				UserPrincipal userToSearch = new UserPrincipal(context) { /*EmailAddress = line*/ GivenName = names[0], Surname = names[1] };

				// Instantiate the searcher to find the user that matches the information set above
				PrincipalSearcher searcher = new PrincipalSearcher(userToSearch);

				// Retrieval could throw an exception from bad data input. Surround in try/catch
				try
				{

					UserPrincipal user = searcher.FindOne() as UserPrincipal;
					if (user != null)
						Users.Add(user);
				}
				catch (Exception e)
				{
					Console.WriteLine("The user with the name of {0} was not found.\r\n{1}", line, e.Message);
				}
			}
			Console.WriteLine("{0} names were found in the input file and {1} users were found and retrieved.", _inputCount, OutputCount);
		}

		private static void OutputResults()
		{
			Console.WriteLine("Outputting Data!");
			using (FileStream fs = File.Create(OutputPath))
			{
				fs.Flush(true);
				fs.Close();
			}

			using (StreamWriter sw = new StreamWriter(OutputPath))
			{
				foreach (UserPrincipal user in Users)
				{
					Console.WriteLine("{0}\t{1}\t{2}", user.SamAccountName, user, user.EmailAddress);
					sw.WriteLine("{0}\t{1}\t{2}", user.SamAccountName, user, user.EmailAddress);
				}
				sw.Flush();
				sw.Close();
			}
		}
	}
}
