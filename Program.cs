using (EventLog eventLog = new EventLog("Application"))
{
    eventLog.Source = "SnaffPoint";
    eventLog.WriteEntry("Request returned with the following status: " + status, EventLogEntryType.Warning, 3001);
}using SearchQueryTool.Helpers;
using SearchQueryTool.Model;
using System;
using System.Collections.Specialized;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;

namespace SnaffPoint
{
    class Program
    {
        private static string PresetPath = "./presets";
        private static int MaxRows = 50;
        private static string SingleQueryText = null;
        private static SearchPresetList _SearchPresets;
        private static string BearerToken = null;
        private static string SPUrl = null;
        private static string OutPath = ".\\Output";
        private static bool isFQL = false;
        private static string RefinementFilters = null;

        private static string FileName()
        {
            string fileName;
            string currDate = DateTime.Now.ToString("ddMMyyyy");
            fileName = currDate + "_SnaffOut.csv";


            return fileName;
        }

        public static void WriteHeadersToCsv()
        {
            string headers = "Preset Name,Title,Author,DocId,Path,FileExtension,Description,ViewsRecent,LastModifiedTime,SiteName,SiteId,SiteDescription\n";
            try
            {
                File.WriteAllText(Path.Combine(OutPath, FileName()), headers);
            }
            catch (Exception ex)
            {
                using (EventLog eventLog = new EventLog("Application"))
                {
                    eventLog.Source = "SnaffPoint";
                    eventLog.WriteEntry("File can not be written to. Check OutPath variable. Error: " + ex.Message, EventLogEntryType.Warning, 3001);
                }
            }
        }


        private static void LoadSearchPresetsFromFolder(string presetFolderPath)
        {
            try
            {
                _SearchPresets = new SearchPresetList(presetFolderPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to read search presets. Error: " + ex.Message);
                using (EventLog eventLog = new EventLog("Application"))
                {
                    eventLog.Source = "SnaffPoint";
                    eventLog.WriteEntry("Failed to read search presets. Error: " + ex.Message, EventLogEntryType.Warning, 3001);
                }
            }
        }

        private static SearchQueryResult StartSearchQueryRequest(SearchQueryRequest request)
        {
            SearchQueryResult searchResults = null;
            try
            {
                HttpRequestResponsePair requestResponsePair = HttpRequestRunner.RunWebRequest(request);
                //Console.WriteLine(requestResponsePair.Item1);
                if (requestResponsePair != null)
                {
                    HttpWebResponse response = requestResponsePair.Item2;
                    if (null != response)
                    {
                        if (!response.StatusCode.Equals(HttpStatusCode.OK))
                        {
                            string status = String.Format("HTTP {0} {1}", (int)response.StatusCode, response.StatusDescription);
                            Console.WriteLine("Request returned with following status: " + status);

                            using (EventLog eventLog = new EventLog("Application"))
                            {
                                eventLog.Source = "SnaffPoint";
                                eventLog.WriteEntry("Request returned with the following status: " + status, EventLogEntryType.Warning, 3001);
                            }
                        }
                    }
                }
                searchResults = GetResultItem(requestResponsePair);

                // success, return the results
                return searchResults;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Request failed with exception: " + ex.Message);
                using (EventLog eventLog = new EventLog("Application"))
                {
                    eventLog.Source = "SnaffPoint";
                    eventLog.WriteEntry("Request failed with exception: " + ex.Message, EventLogEntryType.Warning, 3001);
                }
            }
            return searchResults;
        }

        private static SearchQueryResult GetResultItem(HttpRequestResponsePair requestResponsePair)
        {
            SearchQueryResult searchResults;
            var request = requestResponsePair.Item1;

            using (var response = requestResponsePair.Item2)
            {
                using (var reader = new StreamReader(response.GetResponseStream()))
                {
                    var content = reader.ReadToEnd();
                    NameValueCollection requestHeaders = new NameValueCollection();
                    foreach (var header in request.Headers.AllKeys)
                    {
                        requestHeaders.Add(header, request.Headers[header]);
                    }

                    NameValueCollection responseHeaders = new NameValueCollection();
                    foreach (var header in response.Headers.AllKeys)
                    {
                        responseHeaders.Add(header, response.Headers[header]);
                    }

                    string requestContent = "";
                    if (request.Method == "POST")
                    {
                        requestContent = requestResponsePair.Item3;
                    }

                    searchResults = new SearchQueryResult
                    {
                        RequestUri = request.RequestUri,
                        RequestMethod = request.Method,
                        RequestContent = requestContent,
                        ContentType = response.ContentType,
                        ResponseContent = content,
                        RequestHeaders = requestHeaders,
                        ResponseHeaders = responseHeaders,
                        StatusCode = response.StatusCode,
                        StatusDescription = response.StatusDescription,
                        HttpProtocolVersion = response.ProtocolVersion.ToString()
                    };
                    searchResults.Process();
                    //Console.WriteLine(content);
                }
            }
            return searchResults;
        }

        static void QueryAllPresets()
        {
            LoadSearchPresetsFromFolder(PresetPath);

            if (_SearchPresets.Presets.Count > 0)
            {
                foreach (var preset in _SearchPresets.Presets)
                {
                    //Console.WriteLine("\n" + preset.Name + "\n" + new String('=', preset.Name.Length) + "\n");
                    preset.Request.Token = BearerToken;
                    preset.Request.SharePointSiteUrl = SPUrl;
                    preset.Request.RowLimit = MaxRows;
                    preset.Request.AcceptType = AcceptType.Json;
                    preset.Request.AuthenticationType = AuthenticationType.SPOManagement; // force to JWT auth method
                    //Console.WriteLine("DEBUG - Request: " + preset.Request.ToString());
                    //Console.WriteLine(preset.Response.ToString());
                    SearchQueryResult results = StartSearchQueryRequest(preset.Request);
                    //DisplayResults(results, preset);
                    ConfigureResults(results, preset);
                }
            }
            else
            {
                Console.WriteLine("No presets were found in " + PresetPath);
                using (EventLog eventLog = new EventLog("Application"))
                {
                    eventLog.Source = "SnaffPoint";
                    eventLog.WriteEntry("No presets found in specified path." , EventLogEntryType.Warning, 3001);
                }
            }
        }



        /*
         * Take Results from QueryAllPresets Function,
         * writes name of preset used, title of file, author, full path of file, and last date modified
         * to a csv file
         * 
         */
        private static void ConfigureResults(SearchQueryResult results, SearchPreset preset)
        {


            //File.WriteAllText(Path.Combine(docPath, "SnaffOut.csv"), string.Empty); //clear existing file
   
            if (results != null)
            {
                if (results.PrimaryQueryResult != null)
                {
                    if (results.PrimaryQueryResult.TotalRows > 0)
                    {
                        foreach (ResultItem item in results.PrimaryQueryResult.RelevantResults)
                        {
                            string path = HttpUtility.UrlEncode(item.Path); string site = HttpUtility.UrlEncode(item.SiteName);
                            string entry = preset.Name + "," + item.Title + "." + item.Extension + "," + item.Author + "," + item.DocId + "," + path + "," + item.Extension + "," + item.Description + "," + item.ViewsRecent + "," + item.LastModifiedTime + "," + site + "," + item.SiteId + "," + item.SiteDescription + "\n";
                            //Console.WriteLine(entry);

                            File.AppendAllText(Path.Combine(OutPath, FileName()), entry);
                        }
                    }
                }
                else
                {
                    using (EventLog eventLog = new EventLog("Application"))
                    {
                        eventLog.Source = "SnaffPoint";
                        eventLog.WriteEntry("Found no results.", EventLogEntryType.Warning, 3001);
                    }
                }
            }
            else
            {
                using (EventLog eventLog = new EventLog("Application"))
                {
                    eventLog.Source = "SnaffPoint";
                    eventLog.WriteEntry("Results are null.", EventLogEntryType.Warning, 3001);
                }
            }
        }

        private static void DisplayResults(SearchQueryResult results, SearchPreset preset)
        {
            if (results != null)
            {
                if (results.PrimaryQueryResult != null)
                {
                    Console.WriteLine("Found " + results.PrimaryQueryResult.TotalRows + " results");
                    if (results.PrimaryQueryResult.TotalRows > MaxRows)
                    {
                        Console.WriteLine("Only showing " + MaxRows + " results, though!");
                    }        
                    if (results.PrimaryQueryResult.TotalRows > 0)
                    {
                        foreach (ResultItem item in results.PrimaryQueryResult.RelevantResults)
                        {
                            Console.WriteLine("---");
                            Console.WriteLine(item.Title);
                            Console.WriteLine(item.Path);
                        }
                    }
                }
                else
                {
                    Console.WriteLine("Found no results... maybe the request failed?");
                }
            }
            else
            {
                Console.WriteLine("Result are null ! What happened there?");
            }
        }

        /*private static void DoSingleQuery()
        {
            // preparing the request for you
            SearchQueryRequest request = new SearchQueryRequest
            {
                SharePointSiteUrl = SPUrl,
                AcceptType = AcceptType.Json,
                Token = BearerToken,
                AuthenticationType = AuthenticationType.SPOManagement,
                QueryText = SingleQueryText,
                HttpMethodType = HttpMethodType.Get,
                EnableFql = isFQL,
                RowLimit = MaxRows
            };
            if (RefinementFilters != null)
            {
                request.RefinementFilters = RefinementFilters;
            }
            // DO IT, DO IT, DO IT !
            SearchQueryResult results = StartSearchQueryRequest(request);
            DisplayResults(results, preset);
        }*/

        static void PrintHelp()
        {
            Console.WriteLine(
@"
  .dBBBBP   dBBBBb dBBBBBb     dBBBBP dBBBBP dBBBBBb  dBBBBP dBP dBBBBb dBBBBBBP
  BP           dBP      BB                       dB' dBP.BP         dBP         
  `BBBBb  dBP dBP   dBP BB   dBBBP  dBBBP    dBBBP' dBP.BP dBP dBP dBP   dBP    
     dBP dBP dBP   dBP  BB  dBP    dBP      dBP    dBP.BP dBP dBP dBP   dBP     
dBBBBP' dBP dBP   dBBBBBBB dBP    dBP      dBP    dBBBBP dBP dBP dBP   dBP      

                           https://github.com/nheiniger/snaffpoint
                                               
SnaffPoint, candy finder for SharePoint

Usage: SnaffPoint.exe -u URL -t JWT [OPTIONS]

-h, --help              This is me :)

Mandatory:
-u, --url               SharePoint online URL where you want to search
-t, --token             Bearer token that grants access to said SharePoint

Common options:
-m, --max-rows          Max. number of rows to return per search query (default is 50)

Presets mode (default):
-p, --preset            Path to a folder containing XML search presets (default is ./presets)

Single query mode:
-q, --query             Query search string
-l, --fql               Enables FQL (default is KQL)
-r, --refinement-filter Adds a refinement filter");
        }

        static void Main(string[] args)
        {

            foreach (var entry in args.Select((value, index) => new { index, value }))
            {
                switch (entry.value)
                {
                    // do you want FQL powaa?
                    case "-l":
                    case "--fql":
                        isFQL = true;
                        break;
                    // no need for hundreds of results
                    case "-m":
                    case "--max-rows":
                        if (args[entry.index + 1].StartsWith("-"))
                        {
                            PrintHelp();
                            return;
                        }
                        if (! int.TryParse(args[entry.index + 1], out MaxRows))
                        {
                            PrintHelp();
                            return;
                        }
                        break;
                    // preset path, load presets
                    case "-p":
                    case "--preset":
                        if (args[entry.index + 1].StartsWith("-"))
                        {
                            PrintHelp();
                            return;
                        }
                        PresetPath = args[entry.index + 1];
                        break;
                    // single query
                    case "-q":
                    case "--query":
                        if (args[entry.index + 1].StartsWith("-"))
                        {
                            PrintHelp();
                            return;
                        }
                        SingleQueryText = args[entry.index + 1];
                        break;
                    // fine control is good :)
                    case "-r":
                    case "--refinement-filter":
                        if (args[entry.index + 1].StartsWith("-"))
                        {
                            PrintHelp();
                            return;
                        }
                        RefinementFilters = args[entry.index + 1];
                        break;
                    // Bearer token (JWT)
                    case "-t":
                    case "--token":
                        if (args[entry.index + 1].StartsWith("-"))
                        {
                            PrintHelp();
                            return;
                        }
                        BearerToken = "Bearer " + args[entry.index + 1];
                        break;
                    // SharePoint online URL
                    case "-u":
                    case "--url":
                        if (args[entry.index + 1].StartsWith("-"))
                        {
                            PrintHelp();
                            return;
                        }
                        SPUrl = args[entry.index + 1];
                        break;
                    // send help
                    case "-h":
                    case "--help":
                        PrintHelp();
                        return;
                }
            }

            // did you read the doc?
            if (SPUrl == null || BearerToken == null)
            {
                PrintHelp();
                return;
            }

            // if you specify a query I assume you want an answer, otherwise I have some defaults
            if (SingleQueryText != null)
            {
                //DoSingleQuery();
            }
            else
            {
                WriteHeadersToCsv();
                QueryAllPresets();
            }
        }
    }
}
