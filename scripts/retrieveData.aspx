<%@ Page Language="C#" Debug="true" %>
<%@ import Namespace="System.Net" %>
<%@ import Namespace="System.IO" %>
<%@ Import Namespace="Newtonsoft.Json" %>
<%@ Import Namespace="Newtonsoft.Json.Linq" %>

<script runat="server">
    // root object of our JSON data
    RootObject data;
    // querystring data
    string token;
    string JSONPCallback;
    // set to true for output to be pure JSON and not JSONP which isn't rendered nicely in browser for testing
    bool testing = false;
    
    internal class Deadline {
        public string id { get; set; }
        public string name { get; set; }
        public string author { get; set; }
        public string authorname { get; set; }
        public int studentType { get; set; }
        public string course { get; set; }
        public string description { get; set; }
        public int dueyear { get; set; }
        public int duemonth { get; set; }
        public int dueday { get; set; }
        public string duetime { get; set; }
    }
    
    internal class RootObject {
        public List<Deadline> deadlines { get; set; }
    }
        
    // -------------------------------------------------------------------------- event handlers
    protected void page_load() {
        // grab querystring variables from request
        if (!testing) token = Request.QueryString["token"];
        else token = "aderqmbjm1i8jd0l170a";
        
        JSONPCallback = Request.QueryString["callback"];

        // construct main root JSON object
        data = new RootObject();
        // construct deadline List 
        data.deadlines = new List<Deadline>();    
        
        addD2LData();    

        // write JSONP to response
        StringBuilder sb = new StringBuilder();
        if (!testing) sb.Append(JSONPCallback + "(");
        sb.Append(JsonConvert.SerializeObject(data, Formatting.Indented)); // indentation is just for ease of reading while testing
        if (!testing) sb.Append(");"); 

        Context.Response.Clear();
        Context.Response.ContentType = "application/json";
        Context.Response.Write(sb.ToString());
        Context.Response.End();
    }

    // -------------------------------------------------------------------------- private methods
    private void addD2LData() {
        WebRequest webRequest = WebRequest.Create("https://nscconline.desire2learn.com/d2l/le/calendar/feed/user/feed.ics?token=" + token);
        webRequest.Method = "GET";
        
        // do the request and read the response from the stream
        WebResponse resp  = webRequest.GetResponse();
        StreamReader responseReader = new StreamReader(webRequest.GetResponse().GetResponseStream());
        // dumps the iCal from the response into a string variable
        string icsString = responseReader.ReadToEnd();
        // check if request actually returns something - otherwise an error
        if (icsString.Substring(0, 5) != "BEGIN") {
            return;
        }

        // parse the ics moodle data (iCal format)
        //Response.Write("BEFORE:<br/>" + icsString + "<br/><br/><br/>")
        
        // cleaning ICS string
        // removing carriage returns at the end of multiline fields (description)
        icsString = Regex.Replace(icsString, @"\r\n\t\s", "");
        
        // replacing commas \t, tabs, and semicolons
        icsString = Regex.Replace(icsString, "\\\\,", ",");
        icsString = Regex.Replace(icsString, "\\t", " ");
        icsString = Regex.Replace(icsString, "\\\\;", ";");
        icsString = Regex.Replace(icsString, "\\>", "&gt;");
        icsString = Regex.Replace(icsString, "\\<", "&lt;");

        // grab all event strings from ics file and store in collection
        MatchCollection events = Regex.Matches(icsString, @"BEGIN\:VEVENT([\s\S]*?)END\:VEVENT");

        // only parse/clean iCal file if it contains data
        if (events.Count > 0) {
            foreach (Match currentEvent in events) {
                // isolate strings of iCal
                string nameString = Regex.Match(currentEvent.Value, @"SUMMARY\:(.*\r\n)").Value;
                
                //Response.Write("<br/><br/>*" & currentEvent.Value & "*<br/><br/>")
                string descriptionString = Regex.Match(currentEvent.Value, @"DESCRIPTION\:(.*)CLASS", RegexOptions.Singleline).Value;
                
                string locationString = Regex.Match(currentEvent.Value, @"LOCATION\:(.*\r\n)").Value;
                string dueDatesString = Regex.Match(currentEvent.Value, @"DTSTART\:(.*\r\n)|DTSTART;VALUE=DATE\:(.*\r\n)").Value;
                string uid = Regex.Match(currentEvent.Value, @"UID\:(.*\r\n)").Value;
                uid = Regex.Replace(uid, @"UID\:", "");
                uid = Regex.Replace(uid, "\r\n", "");
                uid = Regex.Replace(uid, "@nscconline.desire2learn.com", "-D2L");

                string name = Regex.Replace(nameString, @"SUMMARY\:", "");
                name = Regex.Replace(name, "\r\n", "");
                string course = "";
                if (locationString == "") course = "D2L : Personal Event";
                else course = Regex.Replace(locationString, @"LOCATION\:", "");
                course = Regex.Replace(course, "\r\n", "");
                
                // clean description rogue characters <br/>
                string description = descriptionString;
                description = Regex.Replace(description, @"DESCRIPTION\:", "");
                description = Regex.Replace(description, "CLASS", "");
                description = Regex.Replace(description, "\r\n ", "");
                description = Regex.Replace(description, "\r\n", "");
                description = Regex.Replace(description, "\\\\n\\\\n\\\\n", "<br/><br/>");
                description = Regex.Replace(description, "\\\\n", "<br/>");
                if (description == "") description = "None provided";
                
                // swap out long ugly internal D2L urls in description with clean links
                MatchEvaluator evaluator = new MatchEvaluator(makeLink);
                description = Regex.Replace(description, @"(http|ftp|https):\/\/([\w\-_]+(?:(?:\.[\w\-_]+)+))([\w\-\.,@?^=%&amp;:/~\+#]*[\w\-\@?^=%&amp;/~\+#])?", evaluator);
                              
                //Response.Write("<br/><br/>*" + description + "*<br/><br/>");

                // extract date from ICS format
                string dueDateString = Regex.Replace(dueDatesString, @"DTSTART\:", "");
                dueDateString = Regex.Replace(dueDateString, @"DTSTART;VALUE=DATE\:", "");
                int year;
                int month;
                int day;
                int hours = 11;
                int minutes = 30;
                if (dueDateString.Length > 10) {
                    // normal dateStamp
                    year = Convert.ToInt32(dueDateString.Substring(0, 4));
                    month = Convert.ToInt32(dueDateString.Substring(4, 2)) - 1;
                    day = Convert.ToInt32(dueDateString.Substring(6, 2));
                    hours = Convert.ToInt32(dueDateString.Substring(9, 2));
                    minutes = Convert.ToInt32(dueDateString.Substring(11, 2));
                } else {
                    year = Convert.ToInt32(dueDateString.Substring(0, 4));
                    month = Convert.ToInt32(dueDateString.Substring(4, 2)) - 1;
                    day = Convert.ToInt32(dueDateString.Substring(6, 2));
                }
                
                // iCal is storing the moodle event time as if we are in Greenwich, UK (UTC time)- we need to convert it back to local time
                DateTime d = new DateTime(year, month + 1, day, hours, minutes, 0, 0, DateTimeKind.Utc).ToLocalTime();
                // update with converted date to localtime
                year = d.Year;
                month = d.Month;
                day = d.Day;
                // formatting time string
                string time = d.ToString("hh:mm");
                
                // construct JSON object
                Deadline newDeadline = new Deadline();
                newDeadline.id = uid;
                newDeadline.name = name;
                newDeadline.author = "D2L";
                newDeadline.authorname = "D2L";
                newDeadline.studentType = 1;
                newDeadline.course = course;
                newDeadline.description = description;
                newDeadline.dueyear = year;
                newDeadline.duemonth = month;
                newDeadline.dueday = day;
                newDeadline.duetime = time;

                data.deadlines.Add(newDeadline);
            }
        }

        //Response.Write("<br/><br/>AFTER:<br/>" + icsString + "<br/><br/><br/>");
    }

    private string makeLink(Match match) {
        string linkString = "<a href='" + match.Value + "' target='sean'>Click Here</a>";
        return linkString;
    }
    
</script>
