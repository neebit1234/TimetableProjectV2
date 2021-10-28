using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Reflection;

namespace TimetableProjectV2
{
    public partial class TimetableForm : Form
    {
        /// <summary>
        /// A program that allow the creation of a timetable that changes based how much work is completed
        /// </summary>
        Data_Structures myStructure = new Data_Structures();//Creating the pri_queue object that the system run on
        List<string> IDList = new List<string>();//List of IDs relating to places in the timetable
        List<Label> labelList = new List<Label>();//A list of labels to add titles to
        Label currentLabel;//Current Label being worked on
        List<Day> publicDateList = new List<Day>();//Public date list to order by date
        int publicClassID;
        int globalCurrentUserID;
        bool teacherStatus;
        int publicClassSize = 30;
        private PrintDocument printDocument1 = new PrintDocument();
        Bitmap memoryImage;
        string path = Path.GetDirectoryName(Application.ExecutablePath);

        public double ConvertDoubleDecimal(decimal decimalVal)
        {
            double doubleVal;

            // Decimal to double conversion
            doubleVal = System.Convert.ToDouble(decimalVal);
            return doubleVal;
        }

        public decimal ConvertDecimalDouble(double doubleVal)
        {
            decimal decimalVal;

            // Double to decimal conversion
            decimalVal = System.Convert.ToDecimal(doubleVal);
            return decimalVal;
        }

        public TimetableForm()//Setup of form
        {
            InitializeComponent();
            printDocument1.PrintPage += new PrintPageEventHandler(printDocument1_PrintPage);
            GetExcelFile();
        }

        public class Data_Structures//Class for anything involved in moving tasks on and off of the priority queue
        {
            private List<Task> queue = new List<Task>();
            private List<Task> taskList = new List<Task>();
            private List<Event> eventList = new List<Event>();
            private List<RepeatedTask> repeatedTaskList = new List<RepeatedTask>();
            private List<Folder> repeatedTaskFolderList = new List<Folder>();
            private List<DateCompleted> DateCompletedList = new List<DateCompleted>();
            private List<User> UserList = new List<User>();
            private List<UserEvent> UserEventList = new List<UserEvent>();
            private List<TeacherSubject> TeacherSubjectList = new List<TeacherSubject>();
            private List<Teacher> TeacherList = new List<Teacher>();
            private List<Class> ClassList = new List<Class>();
            private List<UserClass> UserClassList = new List<UserClass>();
            private List<Subject> SubjectList = new List<Subject>();
            private List<UserSubject> UserSubjectList = new List<UserSubject>();
            private List<ScheduleBlock> ScheduleBlockList = new List<ScheduleBlock>();
            private int currentItem;
            private Nullable<DateTime> termStartDate = null;
            private int itemsPerDay;
            private Nullable<DateTime> schoolStartTime = null;

            public bool IsEmpty()//Check if queue empty
            {
                if (queue.Count == 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }

            public void Fillqueue()
            {
                foreach(Task currentTask in taskList)
                {
                    queue.Add(currentTask);
                }
            }

            public void Enqueue()//Method that adds item to the rear of the queue
            {
                NewTask();
                queue.Add(taskList[taskList.Count - 1]);
            }

            public Task Dequeue()//Method that returns item to the front of the queue
            {
                if (IsEmpty())
                {
                    Console.WriteLine("Queue Empty");
                    return taskList[0];
                }
                else
                {
                    Sort();
                    Task send = queue[0];
                    queue.RemoveAt(0);
                    return send;
                }
            }

            public void Sort()//Puts highest priority item first in queue
            {
                currentItem = 0;
                for (int i = 0; i < queue.Count(); i++)
                {
                    if (queue[i].GettaskPriority() < queue[currentItem].GettaskPriority())
                    {
                        currentItem = i;
                    }
                }
                queue.Insert(0, queue[currentItem]);
                queue.RemoveAt(currentItem + 1);
            }

            public void NewTask()//Adds a new task object to taskList
            {
                taskList.Add(new Task());
            }

            public void NewEvent()
            {
                eventList.Add(new Event());
            }

            public void NewRepeatedTask()
            {
                repeatedTaskList.Add(new RepeatedTask());
            }

            public void NewrepeatedTaskFolder(int ID, string name, int userID)
            {
                repeatedTaskFolderList.Add(new Folder(ID, name, userID));
            }

            public void NewUser(int ID, string name, int password)
            {
                UserList.Add(new User(ID, name, password));
            }

            public void NewTeacher(int ID, string name, int password)
            {
                TeacherList.Add(new Teacher(ID, name, password));
            }

            public void NewSubject(int ID, string name)
            {
                SubjectList.Add(new Subject(ID, name));
            }

            public List<Task> GettaskList() => taskList;

            public List<Event> GeteventList() => eventList;

            public List<RepeatedTask> GetrepeatedTaskList() => repeatedTaskList;

            public List<Folder> GetrepeatedTaskFolderList() => repeatedTaskFolderList;

            public List<DateCompleted> GetDateCompletedList() => DateCompletedList;

            public List<User> GetUserList() => UserList;

            public List<UserEvent> GetUserEventList() => UserEventList;

            public List<TeacherSubject> GetTeacherSubjectList() => TeacherSubjectList;

            public List<Teacher> GetTeacherList() => TeacherList;

            public List<Class> GetClassList() => ClassList;

            public List<UserClass> GetUserClassList() => UserClassList;

            public List<Subject> GetSubjectList() => SubjectList;

            public List<UserSubject> GetUserSubjectList() => UserSubjectList;

            public List<ScheduleBlock> GetScheduleBlockList() => ScheduleBlockList;

            public List<Task> Getqueue() => queue;

            public Nullable<DateTime> GettermStartDate() => termStartDate;

            public void SettermStartDate(DateTime i)
            {
                termStartDate = i;
            }

            public int GetitemsPerDay() => itemsPerDay;

            public void SetitemsPerDay(int i)
            {
                itemsPerDay = i;
            }

            public Nullable<DateTime> GetschoolStartTime() => schoolStartTime;

            public void SetschoolStartTime(DateTime i)
            {
                schoolStartTime = i;
            }
        }

        public class TaskSuperClass
        {
            protected int ID;
            protected string Title;
            protected string Description;
            protected int UserID;

            public int GetID() => ID;
            public string GetTitle() => Title;
            public string GetDescription() => Description;
            public int GetUserID() => UserID;

            public void SetID(int i)
            {
                ID = i;
            }

            public void SetTitle(string i)
            {
                Title = i;
            }

            public void SetDescription(string i)
            {
                Description = i;
            }

            public void SetUserID(int i)
            {
                UserID = i;
            }
        }

        public class Task : TaskSuperClass//Every task is an instance of this task class that holds all the data about a task
        {
            private DateTime taskDeadline;
            private bool taskCompleted;
            private DateTime taskDateCompleted;
            private double taskPriority;

            public Task(int TaskID, string TaskTitle, double Priority, DateTime Deadline, string TaskDescription, bool Completed, DateTime DateCompleted, int userID)//Task object created from database
            {
                ID = TaskID;
                Title = TaskTitle;
                taskPriority = Priority;
                taskDeadline = Deadline;
                Description = TaskDescription;
                taskCompleted = Completed;
                taskDateCompleted = DateCompleted;
                UserID = userID;
            }

            public Task()//Constructor for UI created Task object
            {
                taskCompleted = false;
            }

            public bool GettaskCompleted() => taskCompleted;

            public DateTime GettaskDateCompleted() => taskDateCompleted;

            public double GettaskPriority() => this.taskPriority;

            public DateTime GettaskDeadline() => taskDeadline;

            public void SettaskPriority(double i)
            {
                taskPriority = i;
            }

            public void SettaskCompleted(bool i)
            {
                taskCompleted = i;
            }

            public void SettaskDateCompleted(DateTime i)
            {
                taskDateCompleted = i;
            }

            public void SettaskDeadline(DateTime i)
            {
                taskDeadline = i;
            }
        }

        public class Event : TaskSuperClass//Every event is an instance of this Event class that holds all the data about an event
        {
            private DateTime eventDate;

            public Event(int EventID, string EventTitle, DateTime Date, string EventDescription)//Task object created from database
            {
                ID = EventID;
                Title = EventTitle;
                eventDate = Date;
                Description = EventDescription;
            }

            public Event()//Constructor for blank event, data added through UI
            {

            }

            public DateTime GeteventDate() => eventDate;

            public void SeteventDate(DateTime i)
            {
                eventDate = i;
            }
        }

        public class RepeatedTask : TaskSuperClass
        {
            private int FolderID;

            public RepeatedTask(int RepeatedTaskID, string RepeatedTaskTitle, string RepeatedTaskDescription, int Folder, int userID)
            {
                ID = RepeatedTaskID;
                Title = RepeatedTaskTitle;
                Description = RepeatedTaskDescription;
                FolderID = Folder;
                UserID = userID;
            }

            public RepeatedTask() { }

            public int GetfolderID() => FolderID;

            public void SetfolderID(int i)
            {
                FolderID = i;
            }
        }

        public class Folder//Folder class. Folders connecting repeated tasks to allow the program to switch between topics
        {
            private int folderID;
            private string folderTitle;
            private int UserID;

            public Folder(int ID, string Title, int userID)
            {
                folderID = ID;
                folderTitle = Title;
                UserID = userID;
            }

            public int GetfolderID() => folderID;

            public string GetfolderTitle() => folderTitle;

            public int GetUserID() => UserID;

            public void SetfolderID(int i)
            {
                folderID = i;
            }
        }

        public class DateCompleted
        {
            private int ID;
            private DateTime dateCompleted;
            private bool repeatedTaskCompleted;
            private int taskID;

            public DateCompleted(int DCID, DateTime Date, bool TaskCompleted, int TID)
            {
                ID = DCID;
                dateCompleted = Date;
                repeatedTaskCompleted = TaskCompleted;
                taskID = TID;
            }

            public DateCompleted() { }

            public int GetID() => ID;

            public DateTime GetdateCompleted() => dateCompleted;

            public bool GetrepeatedTaskCompleted() => repeatedTaskCompleted;

            public int GettaskID() => taskID;

            public void SetID(int i)
            {
                ID = i;
            }

            public void SetdateCompleted(DateTime i)
            {
                dateCompleted = i;
            }

            public void SetrepeatedTaskCompleted(bool i)
            {
                repeatedTaskCompleted = i;
            }
        }

        public class User
        {
            protected int ID;
            protected string name;
            protected int password;
            private int tasksPerDay;

            public User(int id, string Name, int Password, int TasksPerDay)
            {
                ID = id;
                name = Name;
                password = Password;
                tasksPerDay = TasksPerDay;
            }

            public User(int id, string Name, int Password)
            {
                ID = id;
                name = Name;
                password = Password;
                tasksPerDay = 3;//Default 3 tasks a day
            }

            public User() { }

            public int GetID() => ID;
            public string Getname() => name;
            public int Getpassword() => password;
            public int GettasksPerDay() => tasksPerDay;

            public void SetID(int i)
            {
                ID = i;
            }

            public void SettasksPerDay(int i)
            {
                tasksPerDay = i;
            }
        }

        public class UserEvent
        {
            private int UserID;
            private int EventID;

            public UserEvent(int userID, int eventID)
            {
                UserID = userID;
                EventID = eventID;
            }

            public int GetUserID() => UserID;

            public int GetEventID() => EventID;
        }

        public class TeacherSubject
        {
            private int TeacherID;
            private int SubjectID;

            public TeacherSubject(int teacherID, int subjectID)
            {
                TeacherID = teacherID;
                SubjectID = subjectID;
            }

            public int GetTeacherID() => TeacherID;

            public int GetSubjectID() => SubjectID;
        }

        public class Teacher : User
        {
            public Teacher(int id, string Name, int Password)
            {
                ID = id;
                name = Name;
                password = Password;
            }
        }

        public class Class
        {
            private int classID;
            private string classTitle;
            private int subjectID;
            private int teacherID;

            public Class(int ID, string Title, int SubjectID, int TeacherID)
            {
                classID = ID;
                classTitle = Title;
                subjectID = SubjectID;
                teacherID = TeacherID;
            }

            public int GetclassID() => classID;

            public string GetclassTitle() => classTitle;

            public int GetsubjectID() => subjectID;

            public int GetteacherID() => teacherID;

            public void SetclassTitle(string i)
            {
                classTitle = i;
            }

            public void SetsubjectID(int i)
            {
                subjectID = i;
            }

            public void SetteacherID(int i)
            {
                teacherID = i;
            }
        }

        public class UserClass//Linked Database Objects
        {
            private int UserID;
            private int ClassID;

            public UserClass(int userID, int classID)
            {
                UserID = userID;
                ClassID = classID;
            }

            public int GetUserID() => UserID;

            public int GetClassID() => ClassID;
        }

        public class Subject
        {
            private int subjectID;
            private string subjectTitle;

            public Subject(int ID, string Title)
            {
                subjectID = ID;
                subjectTitle = Title;
            }

            public int GetsubjectID() => subjectID;

            public string GetsubjectTitle() => subjectTitle;

            public void SetsubjectID(int i)
            {
                subjectID = i;
            }

            public void SetsubjectTitle(string i)
            {
                subjectTitle = i;
            }
        }

        public class UserSubject//Linked Database Objects
        {
            private int UserID;
            private int SubjectID;

            public UserSubject(int userID, int subjectID)
            {
                UserID = userID;
                SubjectID = subjectID;
            }

            public int GetUserID() => UserID;

            public int GetSubjectID() => SubjectID;
        }

        public class Day
        {
            private DateTime date;
            private DayOfWeek day;

            public Day(DateTime thisDate, DayOfWeek thisDay)
            {
                date = thisDate;
                day = thisDay;
            }

            public DateTime Getdate() => date;

            public DayOfWeek Getday() => day;

            public void Setdate(DateTime i)
            {
                date = i;
            }

            public void Setday(DayOfWeek i)
            {
                day = i;
            }
        }

        public class ScheduleBlock
        {
            private int ID;
            private string periodTitle;
            private TimeSpan periodLength;
            private bool schedulable;
            private DateTime startTime;

            public ScheduleBlock(int ID, string Title, TimeSpan Length, bool Schedulable)
            {
                this.ID = ID;
                periodTitle = Title;
                periodLength = Length;
                schedulable = Schedulable;
            }

            public ScheduleBlock(int ID, string Title, TimeSpan Length, bool Schedulable, DateTime StartTime)
            {
                this.ID = ID;
                periodTitle = Title;
                periodLength = Length;
                schedulable = Schedulable;
                startTime = StartTime;
            }

            public int GetID() => ID;
            public string GetperiodTitle() => periodTitle;
            public TimeSpan GetperiodLength() => periodLength;
            public bool Getschedulable() => schedulable;
            public DateTime GetstartTime() => startTime;

            public void SetstartTime(DateTime i)
            {
                startTime = i;
            }
        }

        private void GetExcelFile()//Connecting to the excel file and importing saved events and tasks
        {
            Excel.Application xlTimetable = new Excel.Application();

            if (!File.Exists(path + "\\Timetable-Excel.xls"))
            {
                CreateExcelFile();
            }
            Excel.Workbook TimetableDB = xlTimetable.Workbooks.Open(path + "\\Timetable-Excel.xls");
            Excel._Worksheet Tasks = TimetableDB.Sheets[1];
            Excel._Worksheet Events = TimetableDB.Sheets[2];
            Excel._Worksheet RepeatedTasks = TimetableDB.Sheets[3];
            Excel._Worksheet RepeatedTasksFolders = TimetableDB.Sheets[4];
            Excel._Worksheet DateCompleted = TimetableDB.Sheets[5];
            Excel._Worksheet Users = TimetableDB.Sheets[6];
            Excel._Worksheet UserEvents = TimetableDB.Sheets[7];
            Excel._Worksheet TeacherSubjects = TimetableDB.Sheets[8];
            Excel._Worksheet Teachers = TimetableDB.Sheets[9];
            Excel._Worksheet Classes = TimetableDB.Sheets[10];
            Excel._Worksheet UserClasses = TimetableDB.Sheets[11];
            Excel._Worksheet Subjects = TimetableDB.Sheets[12];
            Excel._Worksheet Details = TimetableDB.Sheets[13];
            Excel._Worksheet UserSubjects = TimetableDB.Sheets[14];
            Excel._Worksheet ScheduleBlocks = TimetableDB.Sheets[15];
            Excel.Range TasksRange = Tasks.UsedRange;
            Excel.Range EventsRange = Events.UsedRange;
            Excel.Range RepeatedTasksRange = RepeatedTasks.UsedRange;
            Excel.Range RepeatedTasksFoldersRange = RepeatedTasksFolders.UsedRange;
            Excel.Range DateCompletedRange = DateCompleted.UsedRange;
            Excel.Range UsersRange = Users.UsedRange;
            Excel.Range UserEventsRange = UserEvents.UsedRange;
            Excel.Range TeacherSubjectsRange = TeacherSubjects.UsedRange;
            Excel.Range TeachersRange = Teachers.UsedRange;
            Excel.Range ClassesRange = Classes.UsedRange;
            Excel.Range UserClassesRange = UserClasses.UsedRange;
            Excel.Range SubjectsRange = Subjects.UsedRange;
            Excel.Range UserSubjectsRange = UserSubjects.UsedRange;
            Excel.Range ScheduleBlocksRange = ScheduleBlocks.UsedRange;

            myStructure.GettaskList().Clear();

            int n = 1;
            foreach (Excel.Range row in TasksRange.Rows)//Get Task Items from excel database
            {
                if (Tasks.Cells[n,1].Value != null)
                {
                    myStructure.GettaskList().Add(new Task(Convert.ToInt32(Tasks.Cells[n, 1].Value), Tasks.Cells[n, 2].Value.ToString(), Tasks.Cells[n, 3].Value, 
                        Tasks.Cells[n, 4].Value, Tasks.Cells[n, 5].Value.ToString(), Tasks.Cells[n, 6].Value, Tasks.Cells[n, 7].Value, Convert.ToInt32(Tasks.Cells[n, 8].Value)));
                    myStructure.Getqueue().Add(myStructure.GettaskList()[myStructure.GettaskList().Count - 1]);
                    n = n + 1;
                }
            }

            myStructure.GeteventList().Clear();

            int i = 1;
            foreach (Excel.Range row in EventsRange.Rows)//Get Event items from excel database
            {
                if (Events.Cells[i, 1].Value != null)
                {
                    myStructure.GeteventList().Add(new Event(Convert.ToInt32(Events.Cells[i, 1].Value), Events.Cells[i, 2].Value.ToString(), Events.Cells[i, 3].Value, Events.Cells[i, 4].Value.ToString()));
                    i = i + 1;
                }
            }

            myStructure.GetrepeatedTaskList().Clear();

            int j = 1;
            foreach (Excel.Range row in RepeatedTasksRange.Rows)//Get Repeated Tasks items from excel database
            {
                if (RepeatedTasks.Cells[j, 1].Value != null)
                {
                    myStructure.GetrepeatedTaskList().Add(new RepeatedTask(Convert.ToInt32(RepeatedTasks.Cells[j, 1].Value), RepeatedTasks.Cells[j, 2].Value.ToString(), RepeatedTasks.Cells[j, 3].Value, Convert.ToInt32(RepeatedTasks.Cells[j, 4].Value), Convert.ToInt32(RepeatedTasks.Cells[j, 5].Value)));
                    j = j + 1;
                }
            }

            myStructure.GetrepeatedTaskFolderList().Clear();

            int a = 1;
            foreach (Excel.Range row in RepeatedTasksFoldersRange.Rows)//Get Repeated Task Folders from excel database
            {
                if (RepeatedTasksFolders.Cells[a, 1].Value != null)
                {
                    myStructure.GetrepeatedTaskFolderList().Add(new Folder(Convert.ToInt32(RepeatedTasksFolders.Cells[a, 1].Value),  RepeatedTasksFolders.Cells[a, 2].Value, Convert.ToInt32(RepeatedTasksFolders.Cells[a, 3].Value)));
                    a = a + 1;
                }
            }

            myStructure.GetDateCompletedList().Clear();

            int b = 1;
            foreach (Excel.Range row in DateCompletedRange.Rows)//Get Dates of repeated tasks from excel database
            {
                if (DateCompleted.Cells[b, 1].Value != null
                    && DateCompleted.Cells[b, 4].Value != 0)
                {
                    myStructure.GetDateCompletedList().Add(new DateCompleted(Convert.ToInt32(DateCompleted.Cells[b, 1].Value),  DateCompleted.Cells[b, 2].Value, DateCompleted.Cells[b, 3].Value, Convert.ToInt32(DateCompleted.Cells[b, 4].Value)));
                    b = b + 1;
                }
            }

            myStructure.GetUserList().Clear();

            int c = 1;
            foreach (Excel.Range row in UsersRange.Rows)//Get user details from excel database
            {
                if (Users.Cells[c, 1].Value != null)
                {
                    myStructure.GetUserList().Add(new User(Convert.ToInt32(Users.Cells[c, 1].Value), Users.Cells[c, 2].Value, Convert.ToInt32(Users.Cells[c, 3].Value), Convert.ToInt32(Users.Cells[c, 4].Value)));
                    c = c + 1;
                }
            }

            myStructure.GetUserEventList().Clear();

            int d = 1;
            foreach (Excel.Range row in UserEventsRange.Rows)//Get UserEvents from excel database
            {
                if (UserEvents.Cells[d, 1].Value != null)
                {
                    myStructure.GetUserEventList().Add(new UserEvent(Convert.ToInt32(UserEvents.Cells[d, 1].Value), Convert.ToInt32(UserEvents.Cells[d, 2].Value)));
                    d = d + 1;
                }
            }

            myStructure.GetTeacherSubjectList().Clear();

            int e = 1;
            foreach (Excel.Range row in TeacherSubjectsRange.Rows)//Get TeacherSubjects from excel database
            {
                if (TeacherSubjects.Cells[e, 1].Value != null)
                {
                    myStructure.GetTeacherSubjectList().Add(new TeacherSubject(Convert.ToInt32(TeacherSubjects.Cells[e, 1].Value), Convert.ToInt32(TeacherSubjects.Cells[e, 2].Value)));
                    e = e + 1;
                }
            }

            myStructure.GetTeacherList().Clear();

            int f = 1;
            foreach (Excel.Range row in TeachersRange.Rows)//Get Teachers from excel database
            {
                if (Teachers.Cells[f, 1].Value != null)
                {
                    myStructure.GetTeacherList().Add(new Teacher(Convert.ToInt32(Teachers.Cells[f, 1].Value), Teachers.Cells[f, 2].Value, Convert.ToInt32(Teachers.Cells[f, 3].Value)));
                    f = f + 1;
                }
            }

            myStructure.GetClassList().Clear();

            int g = 1;
            foreach (Excel.Range row in ClassesRange.Rows)//Get Classes from excel database
            {
                if (Classes.Cells[g, 1].Value != null)
                {
                    myStructure.GetClassList().Add(new Class(Convert.ToInt32(Classes.Cells[g, 1].Value), Classes.Cells[g, 2].Value, Convert.ToInt32(Classes.Cells[g, 3].Value), Convert.ToInt32(Classes.Cells[g, 4].Value)));
                    g = g + 1;
                }
            }

            myStructure.GetUserClassList().Clear();

            int h = 1;
            foreach (Excel.Range row in UserClassesRange.Rows)//Get UserClasses from excel database
            {
                if (UserClasses.Cells[h, 1].Value != null)
                {
                    myStructure.GetUserClassList().Add(new UserClass(Convert.ToInt32(UserClasses.Cells[h, 1].Value), Convert.ToInt32(UserClasses.Cells[h, 2].Value)));
                    h = h + 1;
                }
            }

            myStructure.GetSubjectList().Clear();

            int k = 1;
            foreach (Excel.Range row in SubjectsRange.Rows)//Get Subjects from excel database
            {
                if (Subjects.Cells[k, 1].Value != null)
                {
                    myStructure.GetSubjectList().Add(new Subject(Convert.ToInt32(Subjects.Cells[k, 1].Value), Subjects.Cells[k, 2].Value));
                    k = k + 1;
                }
            }

            myStructure.GetUserSubjectList().Clear();

            int l = 1;
            foreach (Excel.Range row in UserSubjectsRange.Rows)//Get UserSubjects from excel database
            {
                if (UserSubjects.Cells[l, 1].Value != null)
                {
                    myStructure.GetUserSubjectList().Add(new UserSubject(Convert.ToInt32(UserSubjects.Cells[l, 1].Value), Convert.ToInt32(UserSubjects.Cells[l, 2].Value)));
                    l = l + 1;
                }
            }

            myStructure.GetScheduleBlockList().Clear();

            int m = 1;
            foreach (Excel.Range row in ScheduleBlocksRange.Rows)//Get UserSubjects from excel database
            {
                if (ScheduleBlocks.Cells[m, 1].Value != null)
                {
                    string myString = String.Format("{0}", ScheduleBlocks.Cells[m, 3].Text);//Needed as Excel has no time period
                    myStructure.GetScheduleBlockList().Add(new ScheduleBlock(Convert.ToInt32(ScheduleBlocks.Cells[m, 1].Value), ScheduleBlocks.Cells[m, 2].Value, TimeSpan.Parse(myString), ScheduleBlocks.Cells[m, 4].Value, ScheduleBlocks.Cells[m, 5].Value));
                    m = m + 1;
                }
            }

            if (Details.Cells[2, 1].Value != null)
            {
                myStructure.SettermStartDate(Details.Cells[2, 1].Value);
            }
            if (Details.Cells[2, 2].Value != null)
            {
                myStructure.SetschoolStartTime(Details.Cells[2, 2].Value);
            }

            TimetableDB.Save();
            TimetableDB.Close();
            xlTimetable.Quit();

            DateTime thisDate = DateTime.Today;

            foreach (Task task in myStructure.GettaskList())//Set any expired tasks as completed = true
            {
                if (task.GettaskDateCompleted() <= thisDate.AddDays(-1))
                {
                    task.SettaskCompleted(true);
                }
            }

            foreach (DateCompleted currentDateCompleted in myStructure.GetDateCompletedList())//Set any expired repeated tasks as completed = true
            {
                if (currentDateCompleted.GetdateCompleted() <= thisDate.AddDays(-1))
                {
                    currentDateCompleted.SetrepeatedTaskCompleted(true);
                }
            }
        }

        public void SaveExcelFile()//Connecting to the excel file and exporting events and tasks
        {
            Excel.Application xlTimetable = new Excel.Application();
            //Change excel file route
            //Excel.Workbook TimetableDB = xlTimetable.Workbooks.Open(@"C:\Users\joelm_88qx5be\OneDrive\Documents\Computer Science\Non-Exam Assessment\Timetable.xlsx");
            Excel.Workbook TimetableDB = xlTimetable.Workbooks.Open(path + "\\Timetable-Excel.xls");
            Excel._Worksheet Tasks = TimetableDB.Sheets[1];
            Excel._Worksheet Events = TimetableDB.Sheets[2];
            Excel._Worksheet RepeatedTasks = TimetableDB.Sheets[3];
            Excel._Worksheet RepeatedTasksFolders = TimetableDB.Sheets[4];
            Excel._Worksheet DateCompleted = TimetableDB.Sheets[5];
            Excel._Worksheet Users = TimetableDB.Sheets[6];
            Excel._Worksheet UserEvents = TimetableDB.Sheets[7];
            Excel._Worksheet TeacherSubjects = TimetableDB.Sheets[8];
            Excel._Worksheet Teachers = TimetableDB.Sheets[9];
            Excel._Worksheet Classes = TimetableDB.Sheets[10];
            Excel._Worksheet UserClasses = TimetableDB.Sheets[11];
            Excel._Worksheet Subjects = TimetableDB.Sheets[12];
            Excel._Worksheet Details = TimetableDB.Sheets[13];
            Excel._Worksheet UserSubjects = TimetableDB.Sheets[14];
            Excel._Worksheet ScheduleBlocks = TimetableDB.Sheets[15];

            int n = 1;
            foreach (Task task in myStructure.GettaskList())//Save tasks to excel
            {
                if (task.GetID() == 0)
                {
                    int highestID = 0;
                    foreach (Task currentTask in myStructure.GettaskList())
                    {
                        if (currentTask.GetID() > highestID)
                        {
                            highestID = currentTask.GetID();
                        }
                    }
                    task.SetID(highestID + 1);
                }
                Tasks.Cells[n, 1] = task.GetID();
                Tasks.Cells[n, 2] = task.GetTitle();
                Tasks.Cells[n, 3] = task.GettaskPriority();
                Tasks.Cells[n, 4] = task.GettaskDeadline();
                Tasks.Cells[n, 5] = task.GetDescription();
                Tasks.Cells[n, 6] = task.GettaskCompleted();
                Tasks.Cells[n, 7] = task.GettaskDateCompleted();
                Tasks.Cells[n, 8] = task.GetUserID();
                n = n + 1;
            }

            int i = 1;
            foreach (Event Event in myStructure.GeteventList())//Save events to excel
            {
                if(Event.GetID() == 0)
                {
                    int highestID = 0;
                    foreach (Event currentEvent in myStructure.GeteventList())
                    {
                        if(currentEvent.GetID() > highestID)
                        {
                            highestID = currentEvent.GetID();
                        }
                    }
                    Event.SetID(highestID + 1);
                }
                Events.Cells[i, 1] = Event.GetID();
                Events.Cells[i, 2] = Event.GetTitle();
                Events.Cells[i, 3] = Event.GeteventDate();
                Events.Cells[i, 4] = Event.GetDescription();
                i = i + 1;
            }

            int j = 1;
            foreach (RepeatedTask repeatedTask in myStructure.GetrepeatedTaskList())//Save repeated tasks to excel
            {
                if (repeatedTask.GetID() == 0)
                {
                    int highestID = 0;
                    foreach (RepeatedTask currentRepeatedTask in myStructure.GetrepeatedTaskList())
                    {
                        if (currentRepeatedTask.GetID() > highestID)
                        {
                            highestID = currentRepeatedTask.GetID();
                        }
                    }
                    repeatedTask.SetID(highestID + 1);
                }
                RepeatedTasks.Cells[j, 1] = repeatedTask.GetID();
                RepeatedTasks.Cells[j, 2] = repeatedTask.GetTitle();
                RepeatedTasks.Cells[j, 3] = repeatedTask.GetDescription();
                RepeatedTasks.Cells[j, 4] = repeatedTask.GetfolderID();
                RepeatedTasks.Cells[j, 5] = repeatedTask.GetUserID();
                j = j + 1;
            }

            int a = 1;
            foreach (Folder folder in myStructure.GetrepeatedTaskFolderList())//Save folders to excel
            {
                if (folder.GetfolderID() == 0)
                {
                    int highestID = 0;
                    foreach (Folder currentfolder in myStructure.GetrepeatedTaskFolderList())
                    {
                        if (currentfolder.GetfolderID() > highestID)
                        {
                            highestID = currentfolder.GetfolderID();
                        }
                    }
                    folder.SetfolderID(highestID + 1);
                }
                RepeatedTasksFolders.Cells[a, 1] = folder.GetfolderID();
                RepeatedTasksFolders.Cells[a, 2] = folder.GetfolderTitle();
                RepeatedTasksFolders.Cells[a, 3] = folder.GetUserID();
                a = a + 1;
            }

            int b = 1;
            foreach (DateCompleted dateCompleted in myStructure.GetDateCompletedList())//Save dateCompleted tasks to excel
            {
                if (dateCompleted.GetID() == 0)
                {
                    int highestID = 0;
                    foreach (DateCompleted currentDateCompleted in myStructure.GetDateCompletedList())
                    {
                        if (currentDateCompleted.GetID() > highestID)
                        {
                            highestID = currentDateCompleted.GetID();
                        }
                    }
                    dateCompleted.SetID(highestID + 1);
                }
                DateCompleted.Cells[b, 1] = dateCompleted.GetID();
                DateCompleted.Cells[b, 2] = dateCompleted.GetdateCompleted();
                DateCompleted.Cells[b, 3] = dateCompleted.GetrepeatedTaskCompleted();
                DateCompleted.Cells[b, 4] = dateCompleted.GettaskID();
                b = b + 1;
            }

            int c = 1;
            foreach (User user in myStructure.GetUserList())//Save users to excel
            {
                if (user.GetID() == 0)
                {
                    int highestID = 0;
                    foreach (User currentUser in myStructure.GetUserList())
                    {
                        if (currentUser.GetID() > highestID)
                        {
                            highestID = currentUser.GetID();
                        }
                    }
                    user.SetID(highestID + 1);
                }
                Users.Cells[c, 1] = user.GetID();
                Users.Cells[c, 2] = user.Getname();
                Users.Cells[c, 3] = user.Getpassword();
                Users.Cells[c, 4] = user.GettasksPerDay();
                c = c + 1;
            }

            int d = 1;
            foreach (UserEvent userEvent in myStructure.GetUserEventList())//Save userevents to excel
            {
                UserEvents.Cells[d, 1] = userEvent.GetUserID();
                UserEvents.Cells[d, 2] = userEvent.GetEventID();
                d = d + 1;
            }

            int e = 1;
            foreach (TeacherSubject teacherSubject in myStructure.GetTeacherSubjectList())//Save TeacherSubjects to excel
            {
                TeacherSubjects.Cells[e, 1] = teacherSubject.GetTeacherID();
                TeacherSubjects.Cells[e, 2] = teacherSubject.GetSubjectID();
                e = e + 1;
            }

            int f = 1;
            foreach (Teacher teacher in myStructure.GetTeacherList())//Save teachers to excel
            {
                if (teacher.GetID() == 0)
                {
                    int highestID = 0;
                    foreach (Teacher currentTeacher in myStructure.GetTeacherList())
                    {
                        if (currentTeacher.GetID() > highestID)
                        {
                            highestID = currentTeacher.GetID();
                        }
                    }
                    teacher.SetID(highestID + 1);
                }
                Teachers.Cells[f, 1] = teacher.GetID();
                Teachers.Cells[f, 2] = teacher.Getname();
                Teachers.Cells[f, 3] = teacher.Getpassword();
                f = f + 1;
            }

            int g = 1;
            foreach (Class currentClass in myStructure.GetClassList())//Save classes to excel
            {
                Classes.Cells[g, 1] = currentClass.GetclassID();
                Classes.Cells[g, 2] = currentClass.GetclassTitle();
                Classes.Cells[g, 3] = currentClass.GetsubjectID();
                Classes.Cells[g, 4] = currentClass.GetteacherID();
                g = g + 1;
            }

            int h = 1;
            foreach (UserClass userClass in myStructure.GetUserClassList())//Save userclasses to excel
            {
                UserClasses.Cells[h, 1] = userClass.GetUserID();
                UserClasses.Cells[h, 2] = userClass.GetClassID();
                h = h + 1;
            }

            int k = 1;
            foreach (Subject subject in myStructure.GetSubjectList())//Save subjects to excel
            {
                if (subject.GetsubjectID() == 0)
                {
                    int highestID = 0;
                    foreach (Subject currentSubject in myStructure.GetSubjectList())
                    {
                        if (currentSubject.GetsubjectID() > highestID)
                        {
                            highestID = currentSubject.GetsubjectID();
                        }
                    }
                    subject.SetsubjectID(highestID + 1);
                }
                Subjects.Cells[k, 1] = subject.GetsubjectID();
                Subjects.Cells[k, 2] = subject.GetsubjectTitle();
                k = k + 1;
            }

            int l = 1;
            foreach (UserSubject userSubject in myStructure.GetUserSubjectList())//Save userSubjects to excel
            {
                UserSubjects.Cells[l, 1] = userSubject.GetUserID();
                UserSubjects.Cells[l, 2] = userSubject.GetSubjectID();
                l = l + 1;
            }

            int m = 1;
            foreach (ScheduleBlock scheduleBlock in myStructure.GetScheduleBlockList())//Save scheduleBlocks to excel
            {
                ScheduleBlocks.Cells[m, 1] = scheduleBlock.GetID();
                ScheduleBlocks.Cells[m, 2] = scheduleBlock.GetperiodTitle();
                ScheduleBlocks.Cells[m, 3] = scheduleBlock.GetperiodLength().TotalDays;
                ScheduleBlocks.Cells[m, 4] = scheduleBlock.Getschedulable();
                ScheduleBlocks.Cells[m, 5] = scheduleBlock.GetstartTime();
                m = m + 1;
            }

            Details.Cells[2, 1] = myStructure.GettermStartDate();
            Details.Cells[2, 2] = myStructure.GetschoolStartTime();

            TimetableDB.Save();
            TimetableDB.Close();
            xlTimetable.Quit();
        }

        public void CreateExcelFile()
        {
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Please install Microsoft Excel");
                return;
            }

            MessageBox.Show("No Excel file found - Create new file");

            object misValue = System.Reflection.Missing.Value;//Using missing class as a default

            Excel.Workbook TimetableDB = xlApp.Workbooks.Add(misValue);
            Excel.Worksheet xlWorksheet;
            xlWorksheet = (Excel.Worksheet)TimetableDB.Worksheets.get_Item(1);
            xlWorksheet.Name = "ScheduleBlocks";
            Excel.Range cells = xlWorksheet.Columns[3];
            cells.NumberFormat = "hh:mm";
            xlWorksheet = (Excel.Worksheet)TimetableDB.Worksheets.Add();
            xlWorksheet.Name = "UserSubjects";
            xlWorksheet = (Excel.Worksheet)TimetableDB.Worksheets.Add();
            xlWorksheet.Name = "Details";
            xlWorksheet = (Excel.Worksheet)TimetableDB.Worksheets.Add();
            xlWorksheet.Name = "Subjects";
            xlWorksheet = (Excel.Worksheet)TimetableDB.Worksheets.Add();
            xlWorksheet.Name = "UserClasses";
            xlWorksheet = (Excel.Worksheet)TimetableDB.Worksheets.Add();
            xlWorksheet.Name = "Classes";
            xlWorksheet = (Excel.Worksheet)TimetableDB.Worksheets.Add();
            xlWorksheet.Name = "Teachers";
            xlWorksheet = (Excel.Worksheet)TimetableDB.Worksheets.Add();
            xlWorksheet.Name = "TeacherSubjects";
            xlWorksheet = (Excel.Worksheet)TimetableDB.Worksheets.Add();
            xlWorksheet.Name = "UserEvents";
            xlWorksheet = (Excel.Worksheet)TimetableDB.Worksheets.Add();
            xlWorksheet.Name = "Users";
            xlWorksheet = (Excel.Worksheet)TimetableDB.Worksheets.Add();
            xlWorksheet.Name = "DateCompleted";
            xlWorksheet = (Excel.Worksheet)TimetableDB.Worksheets.Add();
            xlWorksheet.Name = "RepeatedTasksFolders";
            xlWorksheet = (Excel.Worksheet)TimetableDB.Worksheets.Add();
            xlWorksheet.Name = "RepeatedTasks";
            xlWorksheet = (Excel.Worksheet)TimetableDB.Worksheets.Add();
            xlWorksheet.Name = "Events";
            xlWorksheet = (Excel.Worksheet)TimetableDB.Worksheets.Add();
            xlWorksheet.Name = "Tasks";

            TimetableDB.SaveAs(path + "\\Timetable-Excel.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, 
                misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            TimetableDB.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorksheet);
            Marshal.ReleaseComObject(TimetableDB);
            Marshal.ReleaseComObject(xlApp);
        }

        public void MakeTimetable()
        {
            int ScheduleBlockCount = 0;
            foreach (ScheduleBlock scheduleBlock in myStructure.GetScheduleBlockList())
            {
                if (scheduleBlock.Getschedulable())
                {
                    ScheduleBlockCount += 1;
                }
            }
            if (!teacherStatus)
            {
                foreach (User user in myStructure.GetUserList())
                {
                    if (user.GetID() == globalCurrentUserID)
                    {
                        AddLabels(user.GettasksPerDay() + ScheduleBlockCount);
                    }
                }
            }
            else
            {
                AddLabels(ScheduleBlockCount);
            }
            FlowLayoutPanel[] flowLayoutPanels = new FlowLayoutPanel[] { Mo1flowLayoutPanel, Tu1flowLayoutPanel, We1flowLayoutPanel, Th1flowLayoutPanel, Fr1flowLayoutPanel, Sa1flowLayoutPanel, Su1flowLayoutPanel,
            Mo2flowLayoutPanel, Tu2flowLayoutPanel, We2flowLayoutPanel, Th2flowLayoutPanel, Fr2flowLayoutPanel, Sa2flowLayoutPanel, Su2flowLayoutPanel,
            Mo3flowLayoutPanel, Tu3flowLayoutPanel, We3flowLayoutPanel, Th3flowLayoutPanel, Fr3flowLayoutPanel, Sa3flowLayoutPanel, Su3flowLayoutPanel,
            Mo4flowLayoutPanel, Tu4flowLayoutPanel, We4flowLayoutPanel, Th4flowLayoutPanel, Fr4flowLayoutPanel, Sa4flowLayoutPanel, Su4flowLayoutPanel};
            //List of every flow layout panel in correct order as thing doesnt choose correct order, dates no display
            List<Label> Titles = new List<Label>();
            foreach (FlowLayoutPanel Panels in flowLayoutPanels)
            {
                if (Panels != flowLayoutPanel1)
                {
                    foreach (Label label in Panels.Controls)
                    {
                        if (Convert.ToInt32(label.Name.Substring(label.Name.Length - 2, 2)) != 00)
                        {
                            Titles.Add(label);
                        }
                    }
                }
            }
            labelList.Clear();
            labelList.AddRange(Titles);
            IDList.Clear();
            foreach (Label label1 in Titles)
            {
                IDList.Add("0");
            }
            DateTime thisDate = DateTime.Today;
            DayOfWeek thisDay = thisDate.DayOfWeek;
            List<Day> dayList = new List<Day>();

            int daysLost = -6 - DaysToInt(thisDay);//Calculating the range of days to show 
            for (int i = 0; i < 28; i++)
            {
                dayList.Add(new Day(thisDate.AddDays(daysLost + i), thisDay));
                publicDateList.Add(new Day(thisDate.AddDays(daysLost + i), thisDay));
            }


            foreach(var flowLayoutPanel in flowLayoutPanels)//Highlight today
            {
                if(flowLayoutPanel.Name.Substring(0,3)== thisDay.ToString().Substring(0,2) + "2")
                {
                    flowLayoutPanel.BackColor = SystemColors.ActiveCaption;
                }
            }

            foreach (var flowLayoutPanel in flowLayoutPanels)//Removes previous text from labels
            {
                var labels = flowLayoutPanel.Controls
                    .OfType<Label>();

                foreach (var label in labels)
                {
                    label.Text = label.Text = "";
                }
            }

            foreach (var flowLayoutPanel in flowLayoutPanels)//Adds the correct date from dateList to label 00
            {
                var labels = flowLayoutPanel.Controls
                    .OfType<Label>()
                    .Where(label => label.Name.EndsWith("label00"));

                foreach (var label in labels)
                {
                    label.Text = dayList[Convert.ToInt32(label.Name.Substring(2, 2))].Getdate().ToString("dd/MM/yyyy");
                    label.Visible = true;
                }
            }

            //Adding classes
            List<Class> scheduledClassList = new List<Class>();
            List<Class> scheduledLabelClassList = new List<Class>();
            foreach (Label label in labelList)
            {
                int index = Convert.ToInt32(label.Name.Substring(2, 2));
                if (label.Text == ""
                    && dayList[index].Getdate().AddDays(1) >= myStructure.GettermStartDate()
                    && label.Name.Substring(0,2) != "Sa"
                    && label.Name.Substring(0,2) != "Su")
                {
                    //Getting Start Time
                    string startTime;
                    if (ScheduleBlockCount > 0)
                    {
                        List<ScheduleBlock> schedulableBlocks = new List<ScheduleBlock>();
                        foreach (ScheduleBlock scheduleBlock in myStructure.GetScheduleBlockList())
                        {
                            if (scheduleBlock.Getschedulable())
                            {
                                schedulableBlocks.Add(scheduleBlock);
                            }
                        }
                        startTime = "N/A";
                        if (Convert.ToInt32(label.Name.Substring(label.Name.Length - 2, 2)) - 1 < schedulableBlocks.Count) 
                        {
                            startTime = schedulableBlocks[Convert.ToInt32(label.Name.Substring(label.Name.Length - 2, 2)) - 1].GetstartTime().ToString("HH:mm");
                        }
                    }
                    else
                    {
                        startTime = DateTime.Now.ToString("HH:mm");
                    }
                    if (startTime == "N/A")//Skip any non-class labels
                    {
                        continue;
                    }
                    foreach (Class @class in myStructure.GetClassList())
                    {
                        //Check no duplicate class
                        bool duplicateClass = false;
                        foreach (Class scheduledClass in scheduledClassList)
                        {
                            if (scheduledClass == @class)
                            {
                                duplicateClass = true; 
                            }
                        }
                        if (duplicateClass)
                        {
                            continue;
                        }
                        //Check no duplicate student
                        bool duplicateStudent = false;
                        foreach (Class scheduledLabelClass in scheduledLabelClassList)
                        {
                            foreach (UserClass userClass in myStructure.GetUserClassList())
                            {
                                foreach (UserClass userClass2 in myStructure.GetUserClassList())
                                {
                                    if (userClass.GetClassID() == scheduledLabelClass.GetclassID()
                                    && userClass2.GetClassID() == @class.GetclassID()
                                    && userClass.GetUserID() == userClass2.GetUserID())
                                    {
                                        duplicateStudent = true;
                                    }
                                }
                            }
                        }
                        if (duplicateStudent)
                        {
                            continue;
                        }
                        //Check no duplicate teacher
                        bool duplicateTeacher = false;
                        foreach (Class scheduledLabelClass in scheduledLabelClassList)
                        {
                            if (scheduledLabelClass.GetteacherID() == @class.GetteacherID()
                                && scheduledLabelClass.GetteacherID() != 0)
                            {
                                duplicateTeacher = true;
                            }
                        }
                        if (duplicateTeacher)
                        {
                            continue;
                        }
                        //Adding classes
                        if (myStructure.GetUserClassList().Count() != 0)
                        {
                            if (!teacherStatus)
                            {
                                UserClass last = myStructure.GetUserClassList().Last();
                                foreach (UserClass userClass in myStructure.GetUserClassList())
                                {
                                    if (userClass.GetUserID() == globalCurrentUserID
                                        && @class.GetclassID() == userClass.GetClassID())
                                    {
                                        label.Visible = true;
                                        label.Text = startTime + " - " + @class.GetclassTitle();
                                        IDList[Titles.IndexOf(label)] = ("C" + @class.GetclassID());
                                        scheduledLabelClassList.Add(@class);
                                        break;
                                    }
                                    else if (userClass == last)
                                    {
                                        scheduledLabelClassList.Add(@class);
                                        break;
                                    }
                                }
                            }
                            else
                            {
                                if (@class.GetteacherID() == globalCurrentUserID)
                                {
                                    label.Visible = true;
                                    label.Text = startTime + " - " + @class.GetclassTitle();
                                    IDList[Titles.IndexOf(label)] = ("C" + @class.GetclassID());
                                    scheduledLabelClassList.Add(@class);
                                }
                                else
                                {
                                    scheduledLabelClassList.Add(@class);
                                }
                            }
                        }
                    }
                    foreach (Class @class in scheduledLabelClassList)
                    {
                        bool Added = false;
                        foreach (Class @class2 in scheduledClassList)
                        {
                            if (@class == @class2)
                            {
                                Added = true;
                            }
                        }
                        if (!Added)
                        {
                            scheduledClassList.Add(@class);
                        }
                    }
                    if (scheduledClassList.Count() >= myStructure.GetClassList().Count())
                    {
                        scheduledClassList.Clear();
                    }
                    if (label.Text == "")
                    {
                        label.Visible = true;
                        label.Text = startTime + " - ";
                    }
                }
                scheduledLabelClassList.Clear();
            }

            if (!teacherStatus)//Only add events, tasks and repeated tasks if not a teacher and in revision mode
            {
                foreach (var flowLayoutPanel in flowLayoutPanels)//Adds events on the correct dates
                {
                    var labels = flowLayoutPanel.Controls.OfType<Label>();

                    //Deleting school labels on weekends
                    if (flowLayoutPanel.Name.Substring(0, 2) == "Sa" ^ flowLayoutPanel.Name.Substring(0, 2) == "Su")
                    {
                        foreach (ScheduleBlock scheduleBlock in myStructure.GetScheduleBlockList())
                        {
                            if (scheduleBlock.Getschedulable())
                            {
                                var labels2 = flowLayoutPanel.Controls.OfType<Label>();
                                IDList.RemoveAt(Titles.IndexOf(labels2.ElementAt(1)));
                                labelList.Remove(labels2.ElementAt(1));
                                Titles.Remove(labels2.ElementAt(1));
                                flowLayoutPanel.Controls.RemoveAt(1);
                            }
                        }
                    }

                    foreach (var label in labels)
                    {
                        if (label.Text == "")
                        {
                            foreach (Event currentEvent in myStructure.GeteventList())
                            {
                                foreach (UserEvent userEvent in myStructure.GetUserEventList())
                                {
                                    if (currentEvent.GeteventDate().ToString("dd/MM/yyyy") == dayList[Convert.ToInt32(label.Name.Substring(2, 2))].Getdate().ToString("dd/MM/yyyy")
                                        && userEvent.GetEventID() == currentEvent.GetID()
                                        && userEvent.GetUserID() == globalCurrentUserID)
                                    {
                                        label.Visible = true;
                                        label.Text = currentEvent.GetTitle();
                                        IDList[Titles.IndexOf(label)] = "E" + currentEvent.GetID();
                                    }
                                }
                            }
                        }
                    }
                }

                if (myStructure.Getqueue().Count() != myStructure.GettaskList().Count())//Fill the queue with tasks from tasklist
                {
                    myStructure.Getqueue().Clear();
                    myStructure.Fillqueue();
                }

                int taskMax = myStructure.Getqueue().Count();
                int taskCount = 1;

                while (taskCount <= taskMax)//Go through every item in queue
                {
                    Task currentTask = myStructure.Dequeue();
                    foreach (Label label in labelList)
                    {
                        if (label.Text == (""))//Set task titles on timetable labels
                        {
                            if (!currentTask.GettaskCompleted() && dayList[Convert.ToInt32(label.Name.Substring(2, 2))].Getdate() >= thisDate
                                && currentTask.GetUserID() == globalCurrentUserID)
                            {
                                label.Visible = true;
                                label.Text = currentTask.GetTitle();
                                currentTask.SettaskDateCompleted(dayList[Convert.ToInt32(label.Name.Substring(2, 2))].Getdate());
                                IDList[Titles.IndexOf(label)] = "T" + currentTask.GetID();
                                break;
                            }
                            else if (currentTask.GettaskDateCompleted() == dayList[Convert.ToInt32(label.Name.Substring(2, 2))].Getdate()
                                && currentTask.GetUserID() == globalCurrentUserID)
                            {
                                label.Visible = true;
                                label.Text = currentTask.GetTitle();
                                IDList[Titles.IndexOf(label)] = "T" + currentTask.GetID();
                                break;
                            }
                        }
                    }
                    taskCount = taskCount + 1;
                }

                List<Folder> personalFolderList = new List<Folder>();//Make lists of folders and repeated tasks with only correct user id
                personalFolderList.Clear();
                foreach (Folder folder in myStructure.GetrepeatedTaskFolderList())
                {
                    if (folder.GetUserID() == globalCurrentUserID)
                    {
                        personalFolderList.Add(folder);
                    }
                }

                List<RepeatedTask> personalRepeatedTaskList = new List<RepeatedTask>();
                personalRepeatedTaskList.Clear();
                foreach (RepeatedTask repeatedTask in myStructure.GetrepeatedTaskList())
                {
                    if (repeatedTask.GetUserID() == globalCurrentUserID)
                    {
                        personalRepeatedTaskList.Add(repeatedTask);
                    }
                }

                int folderMax = personalFolderList.Count();
                int folderCount = 0;
                int repeatedTaskMax = personalRepeatedTaskList.Count();
                int repeatedTaskCount = 0;

                foreach (DateCompleted dateCompleted in myStructure.GetDateCompletedList())//Place outdated repeated tasks
                {
                    bool found = false;
                    foreach (Label label in labelList)
                    {
                        if (found == true)
                        {
                            break;
                        }
                        else if (label.Text == "" && dateCompleted.GetdateCompleted() < DateTime.Today)
                        {
                            if (dateCompleted.GetdateCompleted().ToString("dd/MM/yyyy") == dayList[Convert.ToInt32(label.Name.Substring(2, 2))].Getdate().ToString("dd/MM/yyyy"))
                            {
                                foreach (RepeatedTask repeatedTask in personalRepeatedTaskList)
                                {
                                    if (repeatedTask.GetID() == dateCompleted.GettaskID()
                                        && repeatedTask.GetUserID() == globalCurrentUserID)
                                    {
                                        label.Visible = true;
                                        label.Text = repeatedTask.GetTitle();
                                        IDList[Titles.IndexOf(label)] = "R" + repeatedTask.GetID();
                                        found = true;
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }

                foreach (Label label in labelList)
                {
                    if (folderMax != 0 && repeatedTaskMax != 0)//No divide by 0
                    {
                        if (label.Text == "")//Set repeated task titles on timetable labels
                        {
                            Folder currentFolder = personalFolderList[folderCount % folderMax];

                            List<RepeatedTask> currentRepeatedTaskList = new List<RepeatedTask>();
                            foreach (RepeatedTask repeatedTask in personalRepeatedTaskList)
                            {
                                if (repeatedTask.GetfolderID() == currentFolder.GetfolderID())
                                {
                                    currentRepeatedTaskList.Add(repeatedTask);
                                }
                            }

                            repeatedTaskMax = currentRepeatedTaskList.Count;
                            RepeatedTask currentRepeatedTask = currentRepeatedTaskList[repeatedTaskCount % repeatedTaskMax];

                            DateCompleted currentLinkedDateCompleted = new DateCompleted();
                            foreach (DateCompleted dateCompleted in myStructure.GetDateCompletedList())
                            {
                                if (currentRepeatedTask.GetID() == dateCompleted.GettaskID()
                                    && dateCompleted.GetdateCompleted().ToString("dd/MM/yyyy") == dayList[Convert.ToInt32(label.Name.Substring(2, 2))].Getdate().ToString("dd/MM/yyyy")
                                    && dateCompleted.GetdateCompleted() >= DateTime.Today)
                                {
                                    currentLinkedDateCompleted = dateCompleted;
                                }
                            }

                            if (currentLinkedDateCompleted.GetdateCompleted().ToString("dd/MM/yyyy") == dayList[Convert.ToInt32(label.Name.Substring(2, 2))].Getdate().ToString("dd/MM/yyyy"))
                            {
                                label.Visible = true;
                                label.Text = currentRepeatedTask.GetTitle();
                                IDList[Titles.IndexOf(label)] = "R" + currentRepeatedTask.GetID();
                            }
                            else if (dayList[Convert.ToInt32(label.Name.Substring(2, 2))].Getdate() >= DateTime.Today)
                            {
                                label.Visible = true;
                                label.Text = currentRepeatedTask.GetTitle();
                                IDList[Titles.IndexOf(label)] = "R" + currentRepeatedTask.GetID();
                                int highID = 0;
                                foreach (DateCompleted dateCompleted in myStructure.GetDateCompletedList())
                                {
                                    if (dateCompleted.GetID() > highID)
                                    {
                                        highID = dateCompleted.GetID();
                                    }
                                }
                                myStructure.GetDateCompletedList().Add(new DateCompleted(highID + 1, dayList[Convert.ToInt32(label.Name.Substring(2, 2))].Getdate(), false, currentRepeatedTask.GetID()));
                            }

                            folderCount = folderCount + 1;
                            if (folderCount % folderMax == 0)
                            {
                                repeatedTaskCount = repeatedTaskCount + 1;
                            }
                        }
                    }
                }
            }
            SaveExcelFile();
        }

        public void Edit(Label label)
        {
            currentLabel = label;
            string currentID = IDList[labelList.IndexOf(label)];
            if (currentID.StartsWith("E"))//Events Edits
            {
                foreach(Event currentEvent in myStructure.GeteventList())
                {
                    if(currentEvent.GetID().ToString() == currentID.Substring(1))
                    {
                        TimetableLayoutPanel1.Visible = false;
                        eventEditTitleTextbox.Text = currentEvent.GetTitle();
                        eventEditDateTimePicker.Value = currentEvent.GeteventDate();
                        eventEditDescriptionTextbox.Text = currentEvent.GetDescription();
                        eventEditTitleLabel.Visible = true;
                        eventEditTitleTextbox.Visible = true;
                        eventEditDateLabel.Visible = true;
                        eventEditDateTimePicker.Visible = true;
                        eventEditDescriptionLabel.Visible = true;
                        eventEditDescriptionTextbox.Visible = true;
                        eventSaveEventButton.Visible = true;
                    }
                }
            }
            else if (currentID.StartsWith("T"))//Task Edits
            {
                foreach (Task currentTask in myStructure.GettaskList())
                {
                    if (currentTask.GetID().ToString() == currentID.Substring(1))
                    {
                        TimetableLayoutPanel1.Visible = false;
                        taskEditTitleTextbox.Text = currentTask.GetTitle();
                        taskEditPriorityNumericUpDown.Value = ConvertDecimalDouble(currentTask.GettaskPriority());
                        taskEditDeadlineDateTimePicker.Value = currentTask.GettaskDeadline();
                        taskEditCompletedCheckBox.Checked = currentTask.GettaskCompleted();
                        taskEditDescriptionTextbox.Text = currentTask.GetDescription();
                        taskEditTitleLabel.Visible = true;
                        taskEditTitleTextbox.Visible = true;
                        taskEditPriorityLabel.Visible = true;
                        taskEditPriorityNumericUpDown.Visible = true;
                        taskEditDeadlineLabel.Visible = true;
                        taskEditDeadlineDateTimePicker.Visible = true;
                        taskEditCompletedLabel.Visible = true;
                        taskEditCompletedCheckBox.Visible = true;
                        taskEditDescriptionLabel.Visible = true;
                        taskEditDescriptionTextbox.Visible = true;
                        taskSaveTaskButton.Visible = true;
                    }
                }
            }
            else if (currentID.StartsWith("R"))//Repeated Task Edits
            {
                foreach (DateCompleted currentDateCompleted in myStructure.GetDateCompletedList())
                {
                    if (currentDateCompleted.GettaskID().ToString() == currentID.Substring(1) 
                        && currentDateCompleted.GetdateCompleted().ToString("dd/MM/yyyy") == publicDateList[Convert.ToInt32(label.Name.Substring(2, 2))].Getdate().ToString("dd/MM/yyyy"))
                    {
                        foreach(RepeatedTask currentRepeatedTask in myStructure.GetrepeatedTaskList())
                        {
                            if(currentRepeatedTask.GetID() == currentDateCompleted.GettaskID())
                            {
                                TimetableLayoutPanel1.Visible = false;
                                repeatedTaskEditCompletedCheckBox.Checked = currentDateCompleted.GetrepeatedTaskCompleted();
                                repeatedTaskEditDescriptionTextBox.Text = currentRepeatedTask.GetDescription();
                                repeatedTaskEditTitleTextBox.Text = currentRepeatedTask.GetTitle();
                                repeatedTaskSaveRepeatedTaskButton.Visible = true;
                                repeatedTaskEditCompletedCheckBox.Visible = true;
                                repeatedTaskEditCompletedLabel.Visible = true;
                                repeatedTaskEditDescriptionLabel.Visible = true;
                                repeatedTaskEditDescriptionTextBox.Visible = true;
                                repeatedTaskEditTitleLabel.Visible = true;
                                repeatedTaskEditTitleTextBox.Visible = true;
                            }
                        }
                    }
                }
            }
            else if (currentID.StartsWith("C"))//Class Edits
            {
                foreach (Class currentClass in myStructure.GetClassList())
                {
                    if (currentClass.GetclassID().ToString() == currentID.Substring(1)
                        && !teacherStatus)
                    {
                        TimetableLayoutPanel1.Visible = false;
                        classEditTitleLabel.Visible = true;
                        classEditTitleTextBox.Text = currentClass.GetclassTitle();
                        classEditTitleTextBox.ReadOnly = true;
                        classEditTitleTextBox.Visible = true;
                        classEditSubjectLabel.Visible = true;
                        foreach (Subject subject in myStructure.GetSubjectList())
                        {
                            if (subject.GetsubjectID() == currentClass.GetsubjectID())
                            {
                                classEditSubjectTextBox.Text = subject.GetsubjectTitle();
                            }
                        }
                        classEditSubjectTextBox.ReadOnly = true;
                        classEditSubjectTextBox.Visible = true;
                        classEditTeacherLabel.Visible = true;
                        foreach (Teacher teacher in myStructure.GetTeacherList())
                        {
                            if (teacher.GetID() == currentClass.GetteacherID())
                            {
                                classEditTeacherTextBox.Text = teacher.Getname();
                            }
                        }
                        classEditTeacherTextBox.ReadOnly = true;
                        classEditTeacherTextBox.Visible = true;
                        classEditSaveClassButton.Visible = true;
                        classEditLabelCancelButton.Visible = true;
                    }
                    else if (currentClass.GetclassID().ToString() == currentID.Substring(1)
                        && teacherStatus)
                    {
                        TimetableLayoutPanel1.Visible = false;
                        classEditTitleLabel.Visible = true;
                        classEditTitleTextBox.Text = currentClass.GetclassTitle();
                        classEditTitleTextBox.ReadOnly = false;
                        classEditTitleTextBox.Visible = true;
                        classEditSubjectLabel.Visible = true;
                        classEditSubjectComboBox.Items.Clear();
                        foreach (Subject subject in myStructure.GetSubjectList())
                        {
                            classEditSubjectComboBox.Items.Add(subject.GetsubjectTitle());
                            if (subject.GetsubjectID() == currentClass.GetsubjectID())
                            {
                                classEditSubjectComboBox.Text = subject.GetsubjectTitle();
                            }
                        }
                        classEditSubjectComboBox.Visible = true;
                        classEditTeacherLabel.Visible = true;
                        classEditTeacherComboBox.Items.Clear();
                        foreach (Teacher teacher in myStructure.GetTeacherList())
                        {
                            classEditTeacherComboBox.Items.Add(teacher.Getname());
                            if (teacher.GetID() == currentClass.GetteacherID())
                            {
                                classEditTeacherComboBox.Text = teacher.Getname();
                            }
                        }
                        classEditTeacherComboBox.Visible = true;
                        classEditSaveClassButton.Visible = true;
                        classEditLabelCancelButton.Visible = true;
                    }
                }
            }
        }

        public void SaveEdit(Label label)//Saving an edit after opening the edit menu
        {
            string currentID = IDList[labelList.IndexOf(label)];
            if (currentID.StartsWith("E"))//Saving event edits
            {
                foreach (Event currentEvent in myStructure.GeteventList())
                {
                    if (currentEvent.GetID().ToString() == currentID.Substring(1))
                    {
                        currentEvent.SetTitle(eventEditTitleTextbox.Text);
                        currentEvent.SeteventDate(eventEditDateTimePicker.Value);
                        currentEvent.SetDescription(eventEditDescriptionTextbox.Text);
                    }
                }
            }
            else if (currentID.StartsWith("T"))//Saving task edits
            {
                foreach (Task currentTask in myStructure.GettaskList())
                {
                    if (currentTask.GetID().ToString() == currentID.Substring(1))
                    {
                        currentTask.SetTitle(taskEditTitleTextbox.Text);
                        currentTask.SettaskPriority(ConvertDoubleDecimal(taskEditPriorityNumericUpDown.Value));
                        currentTask.SettaskDeadline(taskEditDeadlineDateTimePicker.Value);
                        currentTask.SettaskCompleted(taskEditCompletedCheckBox.Checked);
                        if (taskEditCompletedCheckBox.Checked == false)
                        {
                            currentTask.SettaskDateCompleted(currentTask.GettaskDateCompleted().AddDays(60));
                        }
                        currentTask.SetDescription(taskEditDescriptionTextbox.Text);
                    }
                }
            }
            else if (currentID.StartsWith("R"))//Saving repeated task edits
            {
                foreach (DateCompleted currentDateCompleted in myStructure.GetDateCompletedList())
                {
                    if (currentDateCompleted.GettaskID().ToString() == currentID.Substring(1) 
                        && currentDateCompleted.GetdateCompleted().ToString("dd/MM/yyyy") == publicDateList[Convert.ToInt32(label.Name.Substring(2, 2))].Getdate().ToString("dd/MM/yyyy"))
                    {
                        foreach (RepeatedTask currentRepeatedTask in myStructure.GetrepeatedTaskList())
                        {
                            if (currentRepeatedTask.GetID() == currentDateCompleted.GettaskID())
                            {
                                currentRepeatedTask.SetTitle(repeatedTaskEditTitleTextBox.Text);
                                currentRepeatedTask.SetDescription(repeatedTaskEditDescriptionTextBox.Text);
                                currentDateCompleted.SetrepeatedTaskCompleted(repeatedTaskEditCompletedCheckBox.Checked);
                                currentDateCompleted.SetdateCompleted(DateTime.Today);
                            }
                        }
                    }
                }
            }
            else if (currentID.StartsWith("C") && teacherStatus)//Saving class edits
            {
                foreach (Class currentClass in myStructure.GetClassList())
                {
                    if (currentClass.GetclassID().ToString() == currentID.Substring(1))
                    {
                        currentClass.SetclassTitle(classEditTitleTextBox.Text);
                        foreach (Subject subject in myStructure.GetSubjectList())
                        {
                            if (subject.GetsubjectTitle() == classEditSubjectComboBox.Text)
                            {
                                currentClass.SetsubjectID(subject.GetsubjectID());
                            }
                        }
                        foreach (Teacher teacher in myStructure.GetTeacherList())
                        {
                            if (teacher.Getname() == classEditTeacherComboBox.Text)
                            {
                                currentClass.SetteacherID(teacher.GetID());
                            }
                        }
                        break;
                    }
                }
            }
        }

        public int DaysToInt(DayOfWeek thisDay)
        {
            switch (thisDay.ToString())
            {
                case "Monday":
                    return 1;
                case "Tuesday":
                    return 2;
                case "Wednesday":
                    return 3;
                case "Thursday":
                    return 4;
                case "Friday":
                    return 5;
                case "Saturday":
                    return 6;
                case "Sunday":
                    return 7;
                default:
                    return 0;
            }
        }

        public Folder CreateFolder(string Name)
        {
            int highestID = 0;
            foreach(Folder folder in myStructure.GetrepeatedTaskFolderList())
            {
                if (folder.GetfolderID() > highestID)
                {
                    highestID = folder.GetfolderID();
                }
                if (folder.GetfolderTitle() == Name
                    && folder.GetUserID() == globalCurrentUserID)
                {
                    return folder;
                }
            }
            myStructure.NewrepeatedTaskFolder(highestID + 1, Name, globalCurrentUserID);
            return myStructure.GetrepeatedTaskFolderList()[myStructure.GetrepeatedTaskFolderList().Count - 1];
        }

        public Subject CreateSubject(string Name)
        {
            int highestID = 0;
            foreach (Subject subject in myStructure.GetSubjectList())
            {
                if (subject.GetsubjectID() > highestID)
                {
                    highestID = subject.GetsubjectID();
                }
                if (subject.GetsubjectTitle() == Name)
                {
                    return subject;
                }
            }
            myStructure.NewSubject(highestID + 1, Name);
            return myStructure.GetSubjectList()[myStructure.GetSubjectList().Count - 1];
        }

        public User CreateUser(string Name, string Password, bool BoolTeacher)//Creates user and teacher accounts
        {
            int encryptedPassword = HashMethod(Password);
            if (!BoolTeacher)
            {
                int highestID = 0;
                foreach (User user in myStructure.GetUserList())
                {
                    if (user.GetID() > highestID)
                    {
                        highestID = user.GetID();
                    }
                    if (user.Getname() == Name && user.Getpassword() == encryptedPassword)
                    {
                        myStructure.SetitemsPerDay(user.GettasksPerDay());
                        return user;
                    }
                }
                myStructure.NewUser(highestID + 1, Name, encryptedPassword);
                return myStructure.GetUserList()[myStructure.GetUserList().Count - 1];
            }
            else
            {
                int highestID = 0;
                foreach (Teacher teacher in myStructure.GetTeacherList())
                {
                    if (teacher.GetID() > highestID)
                    {
                        highestID = teacher.GetID();
                    }
                    if (teacher.Getname() == Name && teacher.Getpassword() == encryptedPassword)
                    {
                        myStructure.SetitemsPerDay(0);
                        return teacher;
                    }
                }
                myStructure.NewTeacher(highestID + 1, Name, encryptedPassword);
                return myStructure.GetTeacherList()[myStructure.GetTeacherList().Count - 1];
            }
        }

        public bool VerifyUser(string Name, string Password, bool BoolTeacher)
        {
            int encryptedPassword = HashMethod(Password);
            if (!BoolTeacher)//Check user accounts
            {
                foreach (User user in myStructure.GetUserList())
                {
                    if (user.Getname() == Name && user.Getpassword() == encryptedPassword)
                    {
                        globalCurrentUserID = user.GetID();
                        teacherStatus = false;
                        myStructure.SetitemsPerDay(user.GettasksPerDay());
                        return true;
                    }
                }
            }
            else if (BoolTeacher)//Check teacher accounts
            {
                foreach (Teacher teacher in myStructure.GetTeacherList())
                {
                    if (teacher.Getname() == Name && teacher.Getpassword() == encryptedPassword)
                    {
                        globalCurrentUserID = teacher.GetID();
                        teacherStatus = true;
                        myStructure.SetitemsPerDay(0);
                        return true;
                    }
                }
            }
            return false;
        }

        public int HashMethod(string Password)//Hashes input string in sections of 4 
        {
            char[] c;
            long multiplier;
            int intLength = Password.Length / 4;
            long sum = 0;
            for (int i = 0; i < intLength; i++)
            {
                c = Password.Substring(i * 4, 4).ToCharArray();
                multiplier = 1;
                for (int j = 0; j< c.Length; j++)
                {
                    sum = sum + c[j] * multiplier;
                    multiplier = multiplier * 256;
                }
            }

            c = Password.Substring(intLength * 4).ToCharArray();
            multiplier = 1;
            for (int j = 0; j < c.Length; j++)
            {
                sum = sum + c[j] * multiplier;
                multiplier = multiplier * 256;
            }

            return Convert.ToInt32(Math.Abs(sum) % 5003);
        }

        private void CaptureScreen()
        {
            Graphics myGraphics = this.CreateGraphics();
            Size s = this.Size;
            memoryImage = new Bitmap(s.Width, s.Height, myGraphics);
            Graphics memoryGraphics = Graphics.FromImage(memoryImage);
            memoryGraphics.CopyFromScreen(this.Location.X, this.Location.Y, 0, 0, s);
        }

        private void printDocument1_PrintPage(System.Object sender,
               System.Drawing.Printing.PrintPageEventArgs e)
        {
            e.Graphics.DrawImage(memoryImage, 0, 0);
        }

        public void EditClass(string classCode)
        {
            foreach (Class currentClass in myStructure.GetClassList())
            {
                if (currentClass.GetclassTitle() == classCode)
                {
                    publicClassID = currentClass.GetclassID();
                }
            }
            classEditClassCodeLabel.Visible = false;
            classEditClassListBox.Visible = false;
            classEditCancelButton.Visible = false;
            classEditUserLabel.Visible = true;
            classEditUserListBox.Items.Clear();
            foreach (User user in myStructure.GetUserList())
            {
                classEditUserListBox.Items.Add(user.Getname());
                foreach (UserClass userClass in myStructure.GetUserClassList())
                {
                    if (userClass.GetUserID() == user.GetID()
                        && userClass.GetClassID() == publicClassID)
                    {
                        classEditUserListBox.SelectedItems.Add(user.Getname());
                    }
                }
            }
            classEditUserListBox.Visible = true;
            classEditUserCancelButton.Visible = true;
            classEditUserSaveUsersButton.Visible = true;
        }

        public void MakeClass()//Automatically putting users in classes
        {
            foreach (UserSubject userSubject in myStructure.GetUserSubjectList())
            {
                bool userEnrolled = false;
                foreach (Class @class in myStructure.GetClassList())
                {
                    if (@class.GetsubjectID() == userSubject.GetSubjectID())
                    {
                        int classSize = 0;
                        foreach (UserClass userClass in myStructure.GetUserClassList())
                        {
                            if (userClass.GetClassID() == @class.GetclassID())
                            {
                                classSize = classSize + 1;
                                if (userClass.GetUserID() == userSubject.GetUserID())
                                {
                                    userEnrolled = true;
                                }
                            }
                        }

                        if (classSize < publicClassSize
                            && !userEnrolled)
                        {
                            myStructure.GetUserClassList().Add(new UserClass(userSubject.GetUserID(), @class.GetclassID()));
                            userEnrolled = true;
                        }
                    }
                }
                if (!userEnrolled)
                {
                    int highestID = 0;
                    int classNumber = 1;
                    foreach (Class @class in myStructure.GetClassList())//Class ID
                    {
                        if (@class.GetclassID() > highestID)
                        {
                            highestID = @class.GetclassID();
                        }
                        if (@class.GetsubjectID() == userSubject.GetSubjectID())
                        {
                            classNumber = classNumber + 1;
                        }
                    }

                    string subjectTitle = "SubjectTitleError";//Default title for a failed title
                    int teacherID = 0;//Default ID 0 no teachers
                    foreach (Subject subject in myStructure.GetSubjectList())//Finding correct subject title and teacher ID
                    {
                        if (subject.GetsubjectID() == userSubject.GetSubjectID())
                        {
                            subjectTitle = subject.GetsubjectTitle();
                            foreach (TeacherSubject teacherSubject in myStructure.GetTeacherSubjectList())
                            {
                                if (teacherSubject.GetSubjectID() == userSubject.GetSubjectID())
                                {
                                    teacherID = teacherSubject.GetTeacherID();
                                }
                            }
                            break;
                        }
                    }
                    myStructure.GetClassList().Add(new Class(highestID + 1, subjectTitle + "Class" + classNumber.ToString(), userSubject.GetSubjectID(), teacherID));
                }
            }
        }

        public void AddLabels(int amount)//Add (amount) of labels to every day panel
        {
            FlowLayoutPanel[] flowLayoutPanels = new FlowLayoutPanel[] { Mo1flowLayoutPanel, Tu1flowLayoutPanel, We1flowLayoutPanel, Th1flowLayoutPanel, Fr1flowLayoutPanel, Sa1flowLayoutPanel, Su1flowLayoutPanel,
            Mo2flowLayoutPanel, Tu2flowLayoutPanel, We2flowLayoutPanel, Th2flowLayoutPanel, Fr2flowLayoutPanel, Sa2flowLayoutPanel, Su2flowLayoutPanel,
            Mo3flowLayoutPanel, Tu3flowLayoutPanel, We3flowLayoutPanel, Th3flowLayoutPanel, Fr3flowLayoutPanel, Sa3flowLayoutPanel, Su3flowLayoutPanel,
            Mo4flowLayoutPanel, Tu4flowLayoutPanel, We4flowLayoutPanel, Th4flowLayoutPanel, Fr4flowLayoutPanel, Sa4flowLayoutPanel, Su4flowLayoutPanel};
            //var flowLayoutPanels = TimetableLayoutPanel1.Controls.OfType<FlowLayoutPanel>();

            int panelCount = 0;
            foreach (var flowLayoutPanel in flowLayoutPanels)
            {
                if (flowLayoutPanel != flowLayoutPanel1)
                {
                    flowLayoutPanel.Controls.Clear();
                    for (int i = 0; i <= amount; i++)
                    {
                        Label newLabel = new Label();
                        newLabel.Click += new EventHandler(DynamicLabel_Click);
                        int singleDigit = panelCount / 10;
                        string panelCountString = panelCount.ToString();
                        if (singleDigit == 0)
                        {
                            panelCountString = "0" + panelCount.ToString();
                        }
                        string labelNum = i.ToString();
                        if (i < 10)
                        {
                            labelNum = "0" + labelNum;
                        }
                        string name = flowLayoutPanel.Name.Substring(0, 2) + panelCountString + "label" + labelNum;
                        newLabel.Name = name;
                        newLabel.Visible = true;
                        flowLayoutPanel.Controls.Add(newLabel);
                    }
                    panelCount += 1;
                }
            }
        }

        private void DynamicLabel_Click(object sender, EventArgs e)
        {
            Label clickedLabel = sender as Label;
            Edit(clickedLabel);
        }

        //Controls for forms
        private void AddTaskButton_Click(object sender, EventArgs e)
        {
            TimetableLayoutPanel1.Visible = false;
            myStructure.Enqueue();
            taskTitleLabel.Visible = true;
            taskDescriptionLabel.Visible = true;
            taskPriorityLabel.Visible = true;
            taskDeadlineLabel.Visible = true;
            taskTitleEnterTextBox.Visible = true;
            taskDescriptionTextBox.Visible = true;
            taskPriorityNumericUpDown.Visible = true;
            taskDeadlineDateTimePicker.Visible = true;
            taskAddTaskButton.Visible = true;
        }

        private void TableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void FlowLayoutPanel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void TitleLabel_Click(object sender, EventArgs e)
        {

        }

        private void TitleEnterTextBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void DescriptionTextBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void PriorityNumericUpDown_ValueChanged(object sender, EventArgs e)
        {

        }

        private void DeadlineDateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void AddTaskButton2_Click(object sender, EventArgs e)
        {
            myStructure.GettaskList()[myStructure.GettaskList().Count - 1].SetTitle(taskTitleEnterTextBox.Text);
            myStructure.GettaskList()[myStructure.GettaskList().Count - 1].SetDescription(taskDescriptionTextBox.Text);
            myStructure.GettaskList()[myStructure.GettaskList().Count - 1].SettaskPriority(ConvertDoubleDecimal(taskPriorityNumericUpDown.Value));
            taskDeadlineDateTimePicker.Format = DateTimePickerFormat.Custom;
            // Display the date as "23/10/2019".  
            taskDeadlineDateTimePicker.CustomFormat = "dd/MM/yyyy";
            myStructure.GettaskList()[myStructure.GettaskList().Count - 1].SettaskDeadline(taskDeadlineDateTimePicker.Value);
            myStructure.GettaskList()[myStructure.GettaskList().Count - 1].SetUserID(globalCurrentUserID);
            taskTitleLabel.Visible = false;
            taskDescriptionLabel.Visible = false;
            taskPriorityLabel.Visible = false;
            taskDeadlineLabel.Visible = false;
            taskTitleEnterTextBox.Visible = false;
            taskDescriptionTextBox.Visible = false;
            taskPriorityNumericUpDown.Visible = false;
            taskDeadlineDateTimePicker.Visible = false;
            taskAddTaskButton.Visible = false;
            TimetableLayoutPanel1.Visible = true;
            MakeTimetable();
            GetExcelFile();
        }

        private void Label1_Click(object sender, EventArgs e)
        {

        }

        private void Timetable_Load(object sender, EventArgs e)
        {
            
        }

        private void M4flowLayoutPanel_Click(object sender, EventArgs e)
        {

        }

        private void eventTitleTextBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void eventDateDateTimePicker_ValueChanged(object sender, EventArgs e)
        {

        }

        private void eventDescriptionTextBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void AddEventButton_Click(object sender, EventArgs e)
        {
            TimetableLayoutPanel1.Visible = false;
            myStructure.NewEvent();
            eventTitleLabel.Visible = true;
            eventDescriptionLabel.Visible = true;
            eventDateLabel.Visible = true;
            eventTitleTextBox.Visible = true;
            eventDescriptionTextBox.Visible = true;
            eventDateDateTimePicker.Visible = true;
            eventAddEventButton.Visible = true;
        }

        private void eventAddEventButton_Click(object sender, EventArgs e)
        {
            if (eventTitleTextBox.Text != ""
                && eventDescriptionTextBox.Text != "")
            {
                myStructure.GeteventList()[myStructure.GeteventList().Count - 1].SetTitle(eventTitleTextBox.Text);
                eventDateDateTimePicker.Format = DateTimePickerFormat.Custom;
                // Display the date as "23/10/2019".  
                eventDateDateTimePicker.CustomFormat = "dd/MM/yyyy";
                myStructure.GeteventList()[myStructure.GeteventList().Count - 1].SeteventDate(eventDateDateTimePicker.Value);
                myStructure.GeteventList()[myStructure.GeteventList().Count - 1].SetDescription(eventDescriptionTextBox.Text);
                foreach (Event Event in myStructure.GeteventList())//Save events to excel
                {
                    if (Event.GetID() == 0)
                    {
                        int highestID = 0;
                        foreach (Event currentEvent in myStructure.GeteventList())
                        {
                            if (currentEvent.GetID() > highestID)
                            {
                                highestID = currentEvent.GetID();
                            }
                        }
                        Event.SetID(highestID + 1);
                    }
                }
                myStructure.GetUserEventList().Add(new UserEvent(globalCurrentUserID, myStructure.GeteventList()[myStructure.GeteventList().Count - 1].GetID()));
                eventTitleLabel.Visible = false;
                eventDescriptionLabel.Visible = false;
                eventDateLabel.Visible = false;
                eventTitleTextBox.Visible = false;
                eventDescriptionTextBox.Visible = false;
                eventDateDateTimePicker.Visible = false;
                eventAddEventButton.Visible = false;
                TimetableLayoutPanel1.Visible = true;
                MakeTimetable();
                GetExcelFile();
            }
            else
            {
                MessageBox.Show("Invalid input");
            }
        }

        private void Mo00label2_Click(object sender, EventArgs e)
        {
            Edit(Mo00label2);
        }

        private void Fr25label2_Click(object sender, EventArgs e)
        {
            Edit(Fr25label2);
        }

        private void eventSaveEventButton_Click(object sender, EventArgs e)
        {
            if (eventEditDescriptionTextbox.Text != ""
                && eventEditTitleTextbox.Text != "")
            {
                SaveEdit(currentLabel);
                MakeTimetable();
                eventEditTitleLabel.Visible = false;
                eventEditTitleTextbox.Visible = false;
                eventEditDateLabel.Visible = false;
                eventEditDateTimePicker.Visible = false;
                eventEditDescriptionLabel.Visible = false;
                eventEditDescriptionTextbox.Visible = false;
                eventSaveEventButton.Visible = false;
                TimetableLayoutPanel1.Visible = true;
            }
            else
            {
                MessageBox.Show("Invalid input");
            }
        }

        private void Mo00label3_Click(object sender, EventArgs e)
        {
            Edit(Mo00label3);
        }

        private void Mo00label4_Click(object sender, EventArgs e)
        {
            Edit(Mo00label4);
        }

        private void We02label2_Click(object sender, EventArgs e)
        {
            Edit(We02label2);
        }

        private void Tu01label3_Click(object sender, EventArgs e)
        {
            Edit(Tu01label3);
        }

        private void Tu01label4_Click(object sender, EventArgs e)
        {
            Edit(Tu01label4);
        }

        private void Tu01label2_Click(object sender, EventArgs e)
        {
            Edit(Tu01label2);
        }

        private void We02label3_Click(object sender, EventArgs e)
        {
            Edit(We02label3);
        }

        private void We02label4_Click(object sender, EventArgs e)
        {
            Edit(We02label4);
        }

        private void Fr04label4_Click(object sender, EventArgs e)
        {
            Edit(Fr04label4);
        }

        private void Sa05label2_Click(object sender, EventArgs e)
        {
            Edit(Sa05label2);
        }

        private void Sa05label3_Click(object sender, EventArgs e)
        {
            Edit(Sa05label3);
        }

        private void Sa05label4_Click(object sender, EventArgs e)
        {
            Edit(Sa05label4);
        }

        private void Th03label2_Click(object sender, EventArgs e)
        {
            Edit(Th03label2);
        }

        private void Su06label3_Click(object sender, EventArgs e)
        {
            Edit(Su06label3);
        }

        private void Su06label4_Click(object sender, EventArgs e)
        {
            Edit(Su06label4);
        }

        private void Mo07label2_Click(object sender, EventArgs e)
        {
            Edit(Mo07label2);
        }

        private void Mo07label3_Click(object sender, EventArgs e)
        {
            Edit(Mo07label3);
        }

        private void Mo07label4_Click(object sender, EventArgs e)
        {
            Edit(Mo07label4);
        }

        private void Th03label4_Click(object sender, EventArgs e)
        {
            Edit(Th03label4);
        }

        private void Su06label2_Click(object sender, EventArgs e)
        {
            Edit(Su06label2);
        }

        private void Th03label3_Click(object sender, EventArgs e)
        {
            Edit(Th03label3);
        }

        private void Tu08label3_Click(object sender, EventArgs e)
        {
            Edit(Tu08label3);
        }

        private void Tu08label4_Click(object sender, EventArgs e)
        {
            Edit(Tu08label4);
        }

        private void We09label2_Click(object sender, EventArgs e)
        {
            Edit(We09label2);
        }

        private void We09label3_Click(object sender, EventArgs e)
        {
            Edit(We09label3);
        }

        private void We09label4_Click(object sender, EventArgs e)
        {
            Edit(We09label4);
        }

        private void Fr04label2_Click(object sender, EventArgs e)
        {
            Edit(Fr04label2);
        }

        private void Tu08label2_Click(object sender, EventArgs e)
        {
            Edit(Tu08label2);
        }

        private void Fr04label3_Click(object sender, EventArgs e)
        {
            Edit(Fr04label3);
        }

        private void Th10label2_Click(object sender, EventArgs e)
        {
            Edit(Th10label2);
        }

        private void Th10label3_Click(object sender, EventArgs e)
        {
            Edit(Th10label3);
        }

        private void Th10label4_Click(object sender, EventArgs e)
        {
            Edit(Th10label4);
        }

        private void Fr11label2_Click(object sender, EventArgs e)
        {
            Edit(Fr11label2);
        }

        private void Fr11label3_Click(object sender, EventArgs e)
        {
            Edit(Fr11label3);
        }

        private void Fr11label4_Click(object sender, EventArgs e)
        {
            Edit(Fr11label4);
        }

        private void Sa12label2_Click(object sender, EventArgs e)
        {
            Edit(Sa12label2);
        }

        private void Sa12label3_Click(object sender, EventArgs e)
        {
            Edit(Sa12label3);
        }

        private void Sa12label4_Click(object sender, EventArgs e)
        {
            Edit(Sa12label4);
        }

        private void Su13label2_Click(object sender, EventArgs e)
        {
            Edit(Su13label2);
        }

        private void Su13label3_Click(object sender, EventArgs e)
        {
            Edit(Su13label3);
        }

        private void Su13label4_Click(object sender, EventArgs e)
        {
            Edit(Su13label4);
        }

        private void Mo14label3_Click(object sender, EventArgs e)
        {
            Edit(Mo14label3);
        }

        private void Mo14label2_Click(object sender, EventArgs e)
        {
            Edit(Mo14label2);
        }

        private void Mo14label4_Click(object sender, EventArgs e)
        {
            Edit(Mo14label4);
        }

        private void Tu15label2_Click(object sender, EventArgs e)
        {
            Edit(Tu15label2);
        }

        private void Tu15label3_Click(object sender, EventArgs e)
        {
            Edit(Tu15label3);
        }

        private void Tu15label4_Click(object sender, EventArgs e)
        {
            Edit(Tu15label4);
        }

        private void We16label2_Click(object sender, EventArgs e)
        {
            Edit(We16label2);
        }

        private void We16label3_Click(object sender, EventArgs e)
        {
            Edit(We16label3);
        }

        private void We16label4_Click(object sender, EventArgs e)
        {
            Edit(We16label4);
        }

        private void Th17label2_Click(object sender, EventArgs e)
        {
            Edit(Th17label2);
        }

        private void Th17label3_Click(object sender, EventArgs e)
        {
            Edit(Th17label3);
        }

        private void Th17label4_Click(object sender, EventArgs e)
        {
            Edit(Th17label4);
        }

        private void Fr18label2_Click(object sender, EventArgs e)
        {
            Edit(Fr18label2);
        }

        private void Fr18label3_Click(object sender, EventArgs e)
        {
            Edit(Fr18label3);
        }

        private void Fr18label4_Click(object sender, EventArgs e)
        {
            Edit(Fr18label4);
        }

        private void Sa19label2_Click(object sender, EventArgs e)
        {
            Edit(Sa19label2);
        }

        private void Sa19label3_Click(object sender, EventArgs e)
        {
            Edit(Sa19label3);
        }

        private void Sa19label4_Click(object sender, EventArgs e)
        {
            Edit(Sa19label4);
        }

        private void Su20label2_Click(object sender, EventArgs e)
        {
            Edit(Su20label2);
        }

        private void Su20label3_Click(object sender, EventArgs e)
        {
            Edit(Su20label3);
        }

        private void Su20label4_Click(object sender, EventArgs e)
        {
            Edit(Su20label4);
        }

        private void Mo21label2_Click(object sender, EventArgs e)
        {
            Edit(Mo21label2);
        }

        private void Mo21label3_Click(object sender, EventArgs e)
        {
            Edit(Mo21label3);
        }

        private void Mo21label4_Click(object sender, EventArgs e)
        {
            Edit(Mo21label4);
        }

        private void Tu22label4_Click(object sender, EventArgs e)
        {
            Edit(Tu22label4);
        }

        private void Tu22label2_Click(object sender, EventArgs e)
        {
            Edit(Tu22label2);
        }

        private void Tu22label3_Click(object sender, EventArgs e)
        {
            Edit(Tu22label3);
        }

        private void We23label2_Click(object sender, EventArgs e)
        {
            Edit(We23label2);
        }

        private void We23label3_Click(object sender, EventArgs e)
        {
            Edit(We23label3);
        }

        private void We23label4_Click(object sender, EventArgs e)
        {
            Edit(We23label4);
        }

        private void Th24label2_Click(object sender, EventArgs e)
        {
            Edit(Th24label2);
        }

        private void Th24label3_Click(object sender, EventArgs e)
        {
            Edit(Th24label3);
        }

        private void Th24label4_Click(object sender, EventArgs e)
        {
            Edit(Th24label4);
        }

        private void Fr25label3_Click(object sender, EventArgs e)
        {
            Edit(Fr25label3);
        }

        private void Fr25label4_Click(object sender, EventArgs e)
        {
            Edit(Fr25label4);
        }

        private void Sa26label2_Click(object sender, EventArgs e)
        {
            Edit(Sa26label2);
        }

        private void Sa26label3_Click(object sender, EventArgs e)
        {
            Edit(Sa26label3);
        }

        private void Sa26label4_Click(object sender, EventArgs e)
        {
            Edit(Sa26label4);
        }

        private void Su27label2_Click(object sender, EventArgs e)
        {
            Edit(Su27label2);
        }

        private void Su27label3_Click(object sender, EventArgs e)
        {
            Edit(Su27label3);
        }

        private void Su27label4_Click(object sender, EventArgs e)
        {
            Edit(Su27label4);
        }

        private void taskSaveTaskButton_Click_1(object sender, EventArgs e)
        {
            if (taskEditDescriptionTextbox.Text != ""
                && taskEditTitleTextbox.Text != "")
            {
                SaveEdit(currentLabel);
                MakeTimetable();
                taskEditTitleLabel.Visible = false;
                taskEditTitleTextbox.Visible = false;
                taskEditPriorityLabel.Visible = false;
                taskEditPriorityNumericUpDown.Visible = false;
                taskEditDeadlineLabel.Visible = false;
                taskEditDeadlineDateTimePicker.Visible = false;
                taskEditCompletedLabel.Visible = false;
                taskEditCompletedCheckBox.Visible = false;
                taskEditDescriptionLabel.Visible = false;
                taskEditDescriptionTextbox.Visible = false;
                taskSaveTaskButton.Visible = false;
                TimetableLayoutPanel1.Visible = true;
            }
            else
            {
                MessageBox.Show("Invalid input");
            }
        }

        private void repeatedTaskAddRepeatedTaskButton_Click(object sender, EventArgs e)
        {
            if (repeatedTaskTitleTextBox.Text != ""
                && repeatedTaskDescriptionTextBox.Text != ""
                && repeatedTaskFolderComboBox.Text != "")
            {
                myStructure.GetrepeatedTaskList()[myStructure.GetrepeatedTaskList().Count - 1].SetTitle(repeatedTaskTitleTextBox.Text);
                myStructure.GetrepeatedTaskList()[myStructure.GetrepeatedTaskList().Count - 1].SetDescription(repeatedTaskDescriptionTextBox.Text);
                myStructure.GetrepeatedTaskList()[myStructure.GetrepeatedTaskList().Count - 1].SetfolderID(CreateFolder(repeatedTaskFolderComboBox.Text).GetfolderID());
                myStructure.GetrepeatedTaskList()[myStructure.GetrepeatedTaskList().Count - 1].SetUserID(globalCurrentUserID);
                repeatedTaskTitleLabel.Visible = false;
                repeatedTaskTitleTextBox.Visible = false;
                repeatedTaskDescriptionLabel.Visible = false;
                repeatedTaskDescriptionTextBox.Visible = false;
                repeatedTaskAddRepeatedTaskButton.Visible = false;
                repeatedTaskFolderComboBox.Visible = false;
                repeatedTaskFolderLabel.Visible = false;
                TimetableLayoutPanel1.Visible = true;
                MakeTimetable();
                GetExcelFile();
            }
            else
            {
                MessageBox.Show("Invalid input");
            }
        }

        private void AddRepeatedTaskButton_Click(object sender, EventArgs e)
        {
            TimetableLayoutPanel1.Visible = false;
            myStructure.NewRepeatedTask();
            repeatedTaskTitleLabel.Visible = true;
            repeatedTaskTitleTextBox.Visible = true;
            repeatedTaskDescriptionLabel.Visible = true;
            repeatedTaskDescriptionTextBox.Visible = true;
            repeatedTaskAddRepeatedTaskButton.Visible = true;
            repeatedTaskFolderComboBox.Items.Clear();
            foreach (Folder folder in myStructure.GetrepeatedTaskFolderList())
            {
                if (folder.GetUserID() == globalCurrentUserID)
                {
                    repeatedTaskFolderComboBox.Items.Add(folder.GetfolderTitle());
                }
            }
            repeatedTaskFolderComboBox.Visible = true;
            repeatedTaskFolderLabel.Visible = true;
        }

        private void repeatedTaskSaveRepeatedTaskButton_Click(object sender, EventArgs e)
        {
            if (repeatedTaskEditDescriptionTextBox.Text != ""
                && repeatedTaskEditTitleTextBox.Text != "")
            {
                SaveEdit(currentLabel);
                MakeTimetable();
                repeatedTaskEditCompletedCheckBox.Visible = false;
                repeatedTaskEditCompletedLabel.Visible = false;
                repeatedTaskEditDescriptionLabel.Visible = false;
                repeatedTaskEditDescriptionTextBox.Visible = false;
                repeatedTaskEditTitleLabel.Visible = false;
                repeatedTaskEditTitleTextBox.Visible = false;
                repeatedTaskSaveRepeatedTaskButton.Visible = false;
                TimetableLayoutPanel1.Visible = true;
            }
            else
            {
                MessageBox.Show("Invalid input");
            }
        }

        private void registerButton_Click(object sender, EventArgs e)
        {
            registerButton.Visible = false;
            loginButton.Visible = false;
            registerNameLabel.Visible = true;
            registerNameTextBox.Visible = true;
            registerPasswordLabel.Visible = true;
            registerPasswordTextBox.Visible = true;
            registerClientLabel.Visible = true;
            registerClientComboBox.Visible = true;
            registerAddUserButton.Visible = true;
            registerCancelButton.Visible = true;
        }

        private void loginButton_Click(object sender, EventArgs e)
        {
            registerButton.Visible = false;
            loginButton.Visible = false;
            loginNameLabel.Visible = true;
            loginNameTextBox.Visible = true;
            loginPasswordLabel.Visible = true;
            loginPasswordTextBox.Visible = true;
            loginClientLabel.Visible = true;
            loginClientComboBox.Visible = true;
            loginLoginButton.Visible = true;
            loginCancelButton.Visible = true;
        }

        private void registerCancelButton_Click(object sender, EventArgs e)
        {
            registerButton.Visible = true;
            loginButton.Visible = true;
            registerNameLabel.Visible = false;
            registerNameTextBox.Visible = false;
            registerNameTextBox.Clear();
            registerPasswordLabel.Visible = false;
            registerPasswordTextBox.Visible = false;
            registerPasswordTextBox.Clear();
            registerClientLabel.Visible = false;
            registerClientComboBox.Visible = false;
            registerAddUserButton.Visible = false;
            registerCancelButton.Visible = false;
        }

        private void loginCancelButton_Click(object sender, EventArgs e)
        {
            registerButton.Visible = true;
            loginButton.Visible = true;
            loginNameLabel.Visible = false;
            loginNameTextBox.Visible = false;
            loginNameTextBox.Clear();
            loginPasswordLabel.Visible = false;
            loginPasswordTextBox.Visible = false;
            loginPasswordTextBox.Clear();
            loginClientLabel.Visible = false;
            loginClientComboBox.Visible = false;
            loginLoginButton.Visible = false;
            loginCancelButton.Visible = false;
        }

        private void registerAddUserButton_Click(object sender, EventArgs e)
        {
            if (registerNameTextBox.Text != "" 
                && registerPasswordTextBox.Text != ""
                && registerClientComboBox.Text != "")
            {
                if (registerClientComboBox.Text == "User")
                {
                    User currentUser = CreateUser(registerNameTextBox.Text, registerPasswordTextBox.Text, false);
                    globalCurrentUserID = currentUser.GetID();
                    myStructure.SetitemsPerDay(currentUser.GettasksPerDay());
                    teacherStatus = false;
                    registerNameLabel.Visible = false;
                    registerNameTextBox.Visible = false;
                    registerNameTextBox.Clear();
                    registerPasswordLabel.Visible = false;
                    registerPasswordTextBox.Visible = false;
                    registerPasswordTextBox.Clear();
                    registerClientLabel.Visible = false;
                    registerClientComboBox.Visible = false;
                    registerAddUserButton.Visible = false;
                    registerCancelButton.Visible = false;
                    EditClassButton.Visible = false;
                    AddClassButton.Visible = false;
                    AddSubjectButton.Visible = false;
                    AddAccountButton.Visible = false;
                    EditTermStartDateButton.Visible = false;
                    EnrolmentButton.Visible = false;
                    SubjectsTeachersButton.Visible = false;
                    userPreferencesButton.Visible = true;
                    AddTaskButton.Visible = true;
                    AddEventButton.Visible = true;
                    AddRepeatedTaskButton.Visible = true;
                    TimetableLayoutPanel1.Visible = true;
                    MakeTimetable();
                }
                else if (registerClientComboBox.Text == "Teacher")
                {
                    globalCurrentUserID = CreateUser(registerNameTextBox.Text, registerPasswordTextBox.Text, true).GetID();
                    myStructure.SetitemsPerDay(0);
                    teacherStatus = true;
                    registerNameLabel.Visible = false;
                    registerNameTextBox.Visible = false;
                    registerNameTextBox.Clear();
                    registerPasswordLabel.Visible = false;
                    registerPasswordTextBox.Visible = false;
                    registerPasswordTextBox.Clear();
                    registerClientLabel.Visible = false;
                    registerClientComboBox.Visible = false;
                    registerAddUserButton.Visible = false;
                    registerCancelButton.Visible = false;
                    if (myStructure.GettermStartDate() == null)
                    {
                        termStartDateLabel.Visible = true;
                        termStartDateDateTimePicker.Visible = true;
                        termStartDateSaveButton.Visible = true;
                        termStartDateCancelButton.Visible = true;
                    }
                    else
                    {
                        EditClassButton.Visible = true;
                        AddClassButton.Visible = true;
                        AddSubjectButton.Visible = true;
                        AddAccountButton.Visible = true;
                        EditTermStartDateButton.Visible = true;
                        EnrolmentButton.Visible = true;
                        SubjectsTeachersButton.Visible = true;
                        userPreferencesButton.Visible = false;
                        AddTaskButton.Visible = false;
                        AddEventButton.Visible = false;
                        AddRepeatedTaskButton.Visible = false;
                        TimetableLayoutPanel1.Visible = true;
                        MakeTimetable();
                    }
                }
            }
            else
            {
                MessageBox.Show("Username or Password Invalid");
            }
        }

        private void loginLoginButton_Click(object sender, EventArgs e)
        {
            bool tempBoolTeacher = false;
            if (loginClientComboBox.Text == "User")
            {
                tempBoolTeacher = false;
            }
            else if (loginClientComboBox.Text == "Teacher")
            {
                tempBoolTeacher = true;
            }
            if (loginNameTextBox.Text != "" 
                && loginPasswordTextBox.Text != "" 
                && VerifyUser(loginNameTextBox.Text, loginPasswordTextBox.Text, tempBoolTeacher))
            {
                loginNameLabel.Visible = false;
                loginNameTextBox.Visible = false;
                loginNameTextBox.Clear();
                loginPasswordLabel.Visible = false;
                loginPasswordTextBox.Visible = false;
                loginPasswordTextBox.Clear();
                loginClientLabel.Visible = false;
                loginClientComboBox.Visible = false;
                loginLoginButton.Visible = false;
                loginCancelButton.Visible = false;
                userPreferencesButton.Visible = true;
                AddTaskButton.Visible = true;
                AddEventButton.Visible = true;
                AddRepeatedTaskButton.Visible = true;
                AddClassButton.Visible = false;
                AddSubjectButton.Visible = false;
                AddAccountButton.Visible = false;
                EditTermStartDateButton.Visible = false;
                EnrolmentButton.Visible = false;
                SubjectsTeachersButton.Visible = false;
                EditClassButton.Visible = false;
                TimetableLayoutPanel1.Visible = true;
                if (teacherStatus)
                {
                    userPreferencesButton.Visible = false;
                    AddTaskButton.Visible = false;
                    AddEventButton.Visible = false;
                    AddRepeatedTaskButton.Visible = false;
                    AddClassButton.Visible = true;
                    AddSubjectButton.Visible = true;
                    AddAccountButton.Visible = true;
                    EditTermStartDateButton.Visible = true;
                    EnrolmentButton.Visible = true;
                    SubjectsTeachersButton.Visible = true;
                    EditClassButton.Visible = true;
                }
                MakeTimetable();
            }
            else
            {
                MessageBox.Show("Username or Password Invalid");
            }
        }

        private void printTimetableButton_Click(object sender, EventArgs e)
        {
            CaptureScreen();
            printDocument1.Print();
        }

        private void AddClassButton_Click(object sender, EventArgs e)
        {
            TimetableLayoutPanel1.Visible = false;
            classAddClassCodeLabel.Visible = true;
            classAddClassCodeTextBox.Visible = true;
            classAddSubjectLabel.Visible = true;
            classAddSubjectComboBox.Items.Clear();
            foreach (Subject subject in myStructure.GetSubjectList())
            {
                classAddSubjectComboBox.Items.Add(subject.GetsubjectTitle());
            }
            classAddSubjectComboBox.Visible = true;
            classAddTeacherLabel.Visible = true;
            classAddTeacherComboBox.Items.Clear();
            foreach (Teacher teacher in myStructure.GetTeacherList())
            {
                classAddTeacherComboBox.Items.Add(teacher.Getname());
            }
            classAddTeacherComboBox.Visible = true;
            classAddClassButton.Visible = true;
            classAddCancelButton.Visible = true;
        }

        private void classAddClassButton_Click(object sender, EventArgs e)//
        {
            if (classAddClassCodeTextBox.Text != ""
                && classAddSubjectComboBox.Text != ""
                && classAddTeacherComboBox.Text != "")
            {
                int tempID = CreateSubject(classAddSubjectComboBox.Text).GetsubjectID();
                int tempTID = 0;
                int highestID = 0;
                foreach (Class currentClass in myStructure.GetClassList())
                {
                    if (currentClass.GetclassID() > highestID)
                    {
                        highestID = currentClass.GetclassID();
                    }
                }
                foreach (Teacher teacher in myStructure.GetTeacherList())
                {
                    if (teacher.Getname() == classAddTeacherComboBox.Text)
                    {
                        tempTID = teacher.GetID();
                    }
                }
                myStructure.GetClassList().Add(new Class(highestID + 1, classAddClassCodeTextBox.Text, tempID, tempTID));
                classAddClassCodeLabel.Visible = false;
                classAddClassCodeTextBox.Visible = false;
                classAddSubjectLabel.Visible = false;
                classAddSubjectComboBox.Visible = false;
                classAddTeacherComboBox.Visible = false;
                classAddTeacherLabel.Visible = false;
                classAddClassButton.Visible = false;
                classAddCancelButton.Visible = false;
                TimetableLayoutPanel1.Visible = true;
                MakeTimetable();
                GetExcelFile();
            }
            else
            {
                MessageBox.Show("Invalid Input");
            }
        }

        private void classEditClassListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (classEditClassListBox.SelectedItem != null)
            {
                EditClass(classEditClassListBox.SelectedItem.ToString());
            }
        }

        private void classEditCancelButton_Click(object sender, EventArgs e)
        {
            classEditClassCodeLabel.Visible = false;
            classEditClassListBox.Visible = false;
            classEditCancelButton.Visible = false;
            TimetableLayoutPanel1.Visible = true;
        }

        private void EditClassButton_Click(object sender, EventArgs e)
        {
            TimetableLayoutPanel1.Visible = false;
            classEditClassCodeLabel.Visible = true;
            classEditClassListBox.Items.Clear();
            foreach (Class currentClass in myStructure.GetClassList())
            {
                classEditClassListBox.Items.Add(currentClass.GetclassTitle());
            }
            classEditClassListBox.Visible = true;
            classEditCancelButton.Visible = true;
        }

        private void classEditUserSaveUsersButton_Click(object sender, EventArgs e)
        {
            foreach (string userName in classEditUserListBox.SelectedItems)
            {
                int tempUserID = 0;
                foreach (User user in myStructure.GetUserList())
                {
                    if (user.Getname() == userName)
                    {
                        tempUserID = user.GetID();
                    }
                }
                bool Added = false;
                foreach (UserClass userClass in myStructure.GetUserClassList())
                {
                    if (userClass.GetUserID() == tempUserID
                        && userClass.GetClassID() == publicClassID)
                    {
                        Added = true;
                    }
                }
                if (!Added)
                {
                    myStructure.GetUserClassList().Add(new UserClass(tempUserID, publicClassID));
                }
                classEditUserLabel.Visible = false;
                classEditUserListBox.Visible = false;
                classEditUserCancelButton.Visible = false;
                classEditUserSaveUsersButton.Visible = false;
                TimetableLayoutPanel1.Visible = true;
                MakeTimetable();
                GetExcelFile();
            }
        }

        private void classEditUserCancelButton_Click(object sender, EventArgs e)
        {
            classEditUserLabel.Visible = false;
            classEditUserListBox.Visible = false;
            classEditUserSaveUsersButton.Visible = false;
            classEditUserCancelButton.Visible = false;
            classEditCancelButton.Visible = true;
            classEditClassCodeLabel.Visible = true;
            classEditClassListBox.Visible = true;
        }

        private void LogoutButton_Click(object sender, EventArgs e)
        {
            globalCurrentUserID = 0;
            teacherStatus = false;
            TimetableLayoutPanel1.Visible = false;
            loginButton.Visible = true;
            registerButton.Visible = true;
        }

        private void AddAccountButton_Click(object sender, EventArgs e)
        {
            TimetableLayoutPanel1.Visible = false;
            accountAddNameLabel.Visible = true;
            accountAddNameTextBox.Visible = true;
            accountAddPasswordLabel.Visible = true;
            accountAddPasswordTextBox.Visible = true;
            accountAddClientLabel.Visible = true;
            accountAddClientComboBox.Visible = true;
            accountAddAddUserButton.Visible = true;
            accountAddCancelButton.Visible = true;
        }

        private void accountAddAddUserButton_Click(object sender, EventArgs e)
        {
            if (accountAddNameTextBox.Text != ""
                && accountAddPasswordTextBox.Text != ""
                && accountAddClientComboBox.Text != "")
            {
                if (accountAddClientComboBox.Text == "User")
                {
                    CreateUser(accountAddNameTextBox.Text, accountAddPasswordTextBox.Text, false).GetID();
                    accountAddNameLabel.Visible = false;
                    accountAddNameTextBox.Visible = false;
                    accountAddNameTextBox.Clear();
                    accountAddPasswordLabel.Visible = false;
                    accountAddPasswordTextBox.Visible = false;
                    accountAddPasswordTextBox.Clear();
                    accountAddClientLabel.Visible = false;
                    accountAddClientComboBox.Visible = false;
                    accountAddAddUserButton.Visible = false;
                    accountAddCancelButton.Visible = false;
                    TimetableLayoutPanel1.Visible = true;
                    MakeTimetable();
                }
                else if (accountAddClientComboBox.Text == "Teacher")
                {
                    CreateUser(accountAddNameTextBox.Text, accountAddPasswordTextBox.Text, true).GetID();
                    accountAddNameLabel.Visible = false;
                    accountAddNameTextBox.Visible = false;
                    accountAddNameTextBox.Clear();
                    accountAddPasswordLabel.Visible = false;
                    accountAddPasswordTextBox.Visible = false;
                    accountAddPasswordTextBox.Clear();
                    accountAddClientLabel.Visible = false;
                    accountAddClientComboBox.Visible = false;
                    accountAddAddUserButton.Visible = false;
                    accountAddCancelButton.Visible = false;
                    TimetableLayoutPanel1.Visible = true;
                    MakeTimetable();
                }
            }
            else
            {
                MessageBox.Show("Username or Password Invalid");
            }
        }

        private void accountAddCancelButton_Click(object sender, EventArgs e)
        {
            TimetableLayoutPanel1.Visible = true;
            accountAddNameLabel.Visible = false;
            accountAddNameTextBox.Visible = false;
            accountAddNameTextBox.Clear();
            accountAddPasswordLabel.Visible = false;
            accountAddPasswordTextBox.Visible = false;
            accountAddPasswordTextBox.Clear();
            accountAddClientLabel.Visible = false;
            accountAddClientComboBox.Visible = false;
            accountAddAddUserButton.Visible = false;
            accountAddCancelButton.Visible = false;
        }

        private void classEditSaveClassButton_Click(object sender, EventArgs e)
        {
            if (classEditTitleTextBox.Text != ""
                && (classEditTeacherTextBox.Text != "" | classEditTeacherComboBox.Text != "")
                && (classEditSubjectTextBox.Text != "" | classEditSubjectComboBox.Text != ""))
            {
                SaveEdit(currentLabel);
                MakeTimetable();
                classEditSubjectComboBox.Visible = false;
                classEditSubjectLabel.Visible = false;
                classEditSubjectTextBox.Visible = false;
                classEditTeacherComboBox.Visible = false;
                classEditTeacherLabel.Visible = false;
                classEditTeacherTextBox.Visible = false;
                classEditTitleLabel.Visible = false;
                classEditTitleTextBox.Visible = false;
                classEditSaveClassButton.Visible = false;
                classEditLabelCancelButton.Visible = false;
                TimetableLayoutPanel1.Visible = true;
            }
            else
            {
                MessageBox.Show("Invalid input");
            }
        }

        private void classEditLabelCancelButton_Click(object sender, EventArgs e)
        {
            classEditSubjectComboBox.Visible = false;
            classEditSubjectLabel.Visible = false;
            classEditSubjectTextBox.Visible = false;
            classEditTeacherComboBox.Visible = false;
            classEditTeacherLabel.Visible = false;
            classEditTeacherTextBox.Visible = false;
            classEditTitleLabel.Visible = false;
            classEditTitleTextBox.Visible = false;
            classEditSaveClassButton.Visible = false;
            classEditLabelCancelButton.Visible = false;
            TimetableLayoutPanel1.Visible = true;
        }

        private void termStartDateSaveButton_Click(object sender, EventArgs e)
        {
            //Saving data
            termStartDateDateTimePicker.Format = DateTimePickerFormat.Custom;
            // Display the date as "23/10/2019".  
            termStartDateDateTimePicker.CustomFormat = "dd/MM/yyyy";
            myStructure.SettermStartDate(termStartDateDateTimePicker.Value);
            propertiesDayStartTimedateTimePicker.Format = DateTimePickerFormat.Custom;
            // Display the time as "10:15".  
            propertiesDayStartTimedateTimePicker.CustomFormat = "HH:mm";
            myStructure.SetschoolStartTime(propertiesDayStartTimedateTimePicker.Value);
            //Visability of components
            termStartDateLabel.Visible = false;
            termStartDateDateTimePicker.Visible = false;
            termStartDateSaveButton.Visible = false;
            termStartDateCancelButton.Visible = false;
            propertiesDayStartTimeLabel.Visible = false;
            propertiesDayStartTimedateTimePicker.Visible = false;
            propertiesTimetableTableLayoutPanel.Visible = false;
            propertiesPeriodTitleLabel.Visible = false;
            propertiesPeriodTitleTextBox.Visible = false;
            propertiesPeriodLengthLabel.Visible = false;
            propertiesPeriodLengthNumericUpDown.Visible = false;
            propertiesSchedulableLabel.Visible = false;
            propertiesSchedulableCheckBox.Visible = false;
            propertiesAddBlockButton.Visible = false;
            propertiesDeleteBlockButton.Visible = false;
            EditClassButton.Visible = true;
            AddClassButton.Visible = true;
            AddSubjectButton.Visible = true;
            AddAccountButton.Visible = true;
            EditTermStartDateButton.Visible = true;
            EnrolmentButton.Visible = true;
            SubjectsTeachersButton.Visible = true;
            userPreferencesButton.Visible = false;
            AddTaskButton.Visible = false;
            AddEventButton.Visible = false;
            AddRepeatedTaskButton.Visible = false;
            TimetableLayoutPanel1.Visible = true;
            MakeTimetable();
        }

        private void termStartDateCancelButton_Click(object sender, EventArgs e)
        {
            termStartDateLabel.Visible = false;
            termStartDateDateTimePicker.Visible = false;
            termStartDateSaveButton.Visible = false;
            termStartDateCancelButton.Visible = false;
            propertiesDayStartTimeLabel.Visible = false;
            propertiesDayStartTimedateTimePicker.Visible = false;
            propertiesTimetableTableLayoutPanel.Visible = false;
            propertiesPeriodTitleLabel.Visible = false;
            propertiesPeriodTitleTextBox.Visible = false;
            propertiesPeriodLengthLabel.Visible = false;
            propertiesPeriodLengthNumericUpDown.Visible = false;
            propertiesSchedulableLabel.Visible = false;
            propertiesSchedulableCheckBox.Visible = false;
            propertiesAddBlockButton.Visible = false;
            propertiesDeleteBlockButton.Visible = false;
            if (myStructure.GettermStartDate() == null)
            {
                myStructure.GetUserList().RemoveAt(myStructure.GetUserList().Count() - 1);
                teacherStatus = false;
                registerNameLabel.Visible = true;
                registerNameTextBox.Visible = true;
                registerNameTextBox.Clear();
                registerPasswordLabel.Visible = true;
                registerPasswordTextBox.Visible = true;
                registerPasswordTextBox.Clear();
                registerClientLabel.Visible = true;
                registerClientComboBox.Visible = true;
                registerAddUserButton.Visible = true;
                registerCancelButton.Visible = true;
            }
            else
            {
                TimetableLayoutPanel1.Visible = true;
            }
        }

        private void EditTermStartDateButton_Click(object sender, EventArgs e)
        {
            TimetableLayoutPanel1.Visible = false;
            termStartDateLabel.Visible = true;
            termStartDateDateTimePicker.Format = DateTimePickerFormat.Custom;
            termStartDateDateTimePicker.CustomFormat = "dd/MM/yyyy";
            DateTime updatedTime = myStructure.GettermStartDate().GetValueOrDefault(DateTime.Now);
            termStartDateDateTimePicker.Value = updatedTime;
            termStartDateDateTimePicker.Visible = true;
            propertiesDayStartTimeLabel.Visible = true;
            propertiesDayStartTimedateTimePicker.Format = DateTimePickerFormat.Custom;
            propertiesDayStartTimedateTimePicker.CustomFormat = "HH:mm";
            updatedTime = myStructure.GetschoolStartTime().GetValueOrDefault(DateTime.Now);
            propertiesDayStartTimedateTimePicker.Value = updatedTime;
            propertiesDayStartTimedateTimePicker.Visible = true;
            //Add Scheduled Blocks
            propertiesTimetableTableLayoutPanel.Controls.Clear();
            propertiesTimetableTableLayoutPanel.ColumnCount = 0;
            TimeSpan timeCount = new TimeSpan();
            DateTime startTime = myStructure.GetschoolStartTime().GetValueOrDefault(DateTime.Now);
            foreach (ScheduleBlock scheduleBlock in myStructure.GetScheduleBlockList())
            {
                Label label = new Label();
                string Time = (startTime + timeCount).ToString("HH:mm") + " - " + (startTime + timeCount).Add(scheduleBlock.GetperiodLength()).ToString("HH:mm");
                timeCount = timeCount.Add(scheduleBlock.GetperiodLength());
                label.Text = scheduleBlock.GetperiodTitle() + "\n" + Time;
                label.AutoSize = true;
                propertiesTimetableTableLayoutPanel.ColumnCount += 1;
                propertiesTimetableTableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
                propertiesTimetableTableLayoutPanel.Controls.Add(label);
            }
            propertiesTimetableTableLayoutPanel.Visible = true;
            propertiesPeriodTitleLabel.Visible = true;
            propertiesPeriodTitleTextBox.Visible = true;
            propertiesPeriodLengthLabel.Visible = true;
            propertiesPeriodLengthNumericUpDown.Visible = true;
            propertiesSchedulableLabel.Visible = true;
            propertiesSchedulableCheckBox.Visible = true;
            propertiesAddBlockButton.Visible = true;
            propertiesDeleteBlockButton.Visible = true;
            termStartDateSaveButton.Visible = true;
            termStartDateCancelButton.Visible = true;
        }

        private void classAddCancelButton_Click(object sender, EventArgs e)
        {
            classAddCancelButton.Visible = false;
            classAddClassButton.Visible = false;
            classAddClassCodeLabel.Visible = false;
            classAddClassCodeTextBox.Visible = false;
            classAddSubjectComboBox.Visible = false;
            classAddTeacherComboBox.Visible = false;
            classAddSubjectLabel.Visible = false;
            classAddTeacherLabel.Visible = false;
            TimetableLayoutPanel1.Visible = true;
        }

        private void EnrolmentButton_Click(object sender, EventArgs e)
        {
            TimetableLayoutPanel1.Visible = false;
            subjectUserEditSubjectListBox.Items.Clear();
            foreach (Subject subject in myStructure.GetSubjectList())
            {
                subjectUserEditSubjectListBox.Items.Add(subject.GetsubjectTitle());
            }
            subjectUserEditSubjectListBox.Visible = true;
            subjectUserEditUserCheckedListBox.Visible = true;
            subjectUserSaveEnrolmentButton.Visible = true;
            subjectUserMainMenuButton.Visible = true;
        }

        private void subjectUserEditSubjectListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            foreach (Subject subject in myStructure.GetSubjectList())
            {
                if (subject.GetsubjectTitle() == subjectUserEditSubjectListBox.SelectedItem.ToString())
                {
                    subjectUserEditUserCheckedListBox.Items.Clear();
                    foreach (User user in myStructure.GetUserList())
                    {
                        bool Added = false;
                        foreach (UserSubject userSubject in myStructure.GetUserSubjectList())
                        {
                            if (subject.GetsubjectID() == userSubject.GetSubjectID()
                                && user.GetID() == userSubject.GetUserID())
                            {
                                subjectUserEditUserCheckedListBox.Items.Add(user.Getname(), true);
                                Added = true;
                            }
                        }
                        if (!Added)
                        {
                            subjectUserEditUserCheckedListBox.Items.Add(user.Getname(), false);
                        }
                    }
                    subjectUserEditUserCheckedListBox.Visible = true;
                    break;
                }
            }
        }

        private void subjectUserSaveEnrolmentButton_Click(object sender, EventArgs e)
        {
            bool Added = false;
            foreach (string checkedName in subjectUserEditUserCheckedListBox.CheckedItems)
            {
                foreach (User user in myStructure.GetUserList())
                {
                    Added = false;
                    if (user.Getname() == checkedName)
                    {
                        foreach (UserSubject userSubject in myStructure.GetUserSubjectList())
                        {
                            foreach (Subject subject in myStructure.GetSubjectList())
                            {
                                if (subject.GetsubjectID() == userSubject.GetSubjectID()
                                    && user.GetID() == userSubject.GetUserID()
                                    && subject.GetsubjectTitle() == subjectUserEditSubjectListBox.SelectedItem.ToString())
                                {
                                    Added = true;
                                    break;
                                }
                            }
                        }
                        break;
                    }
                }
                if (!Added)
                {
                    foreach (User user in myStructure.GetUserList())
                    {
                        if (user.Getname() == checkedName)
                        {
                            foreach (Subject subject in myStructure.GetSubjectList())
                            {
                                if (subject.GetsubjectTitle() == subjectUserEditSubjectListBox.SelectedItem.ToString())
                                {
                                    myStructure.GetUserSubjectList().Add(new UserSubject(user.GetID(), subject.GetsubjectID()));
                                    break;
                                }
                            }
                            break;
                        }
                    }
                }
            }
        }

        private void subjectUserMainMenuButton_Click(object sender, EventArgs e)
        {
            MakeClass();
            MakeTimetable();
            subjectUserEditSubjectListBox.Visible = false;
            subjectUserEditUserCheckedListBox.Visible = false;
            subjectUserSaveEnrolmentButton.Visible = false;
            TimetableLayoutPanel1.Visible = true;
            subjectUserMainMenuButton.Visible = false;
        }

        private void subjectsTeachersButton_Click(object sender, EventArgs e)
        {
            TimetableLayoutPanel1.Visible = false;
            teacherSubjectEditSubjectListBox.Items.Clear();
            foreach (Subject subject in myStructure.GetSubjectList())
            {
                teacherSubjectEditSubjectListBox.Items.Add(subject.GetsubjectTitle());
            }
            teacherSubjectEditSubjectListBox.Visible = true;
            teacherSubjectEditTeacherCheckedListBox.Items.Clear();
            teacherSubjectEditTeacherCheckedListBox.Visible = true;
            teacherSubjectSaveSpecialismButton.Visible = true;
            teacherSubjectMainMenuButton.Visible = true;
        }

        private void teacherSubjectEditSubjectListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            foreach (Subject subject in myStructure.GetSubjectList())
            {
                if (subject.GetsubjectTitle() == teacherSubjectEditSubjectListBox.SelectedItem.ToString())
                {
                    teacherSubjectEditTeacherCheckedListBox.Items.Clear();
                    foreach (Teacher teacher in myStructure.GetTeacherList())
                    {
                        bool Added = false;
                        foreach (TeacherSubject teacherSubject in myStructure.GetTeacherSubjectList())
                        {
                            if (subject.GetsubjectID() == teacherSubject.GetSubjectID()
                                && teacher.GetID() == teacherSubject.GetTeacherID())
                            {
                                teacherSubjectEditTeacherCheckedListBox.Items.Add(teacher.Getname(), true);
                                Added = true;
                            }
                        }
                        if (!Added)
                        {
                            teacherSubjectEditTeacherCheckedListBox.Items.Add(teacher.Getname(), false);
                        }
                    }
                    teacherSubjectEditTeacherCheckedListBox.Visible = true;
                    break;
                }
            }
        }

        private void teacherSubjectSaveSpecialismButton_Click(object sender, EventArgs e)
        {
            bool Added = false;
            foreach (string checkedName in teacherSubjectEditTeacherCheckedListBox.CheckedItems)
            {
                foreach (Teacher teacher in myStructure.GetTeacherList())
                {
                    Added = false;
                    if (teacher.Getname() == checkedName)
                    {
                        foreach (TeacherSubject teacherSubject in myStructure.GetTeacherSubjectList())
                        {
                            foreach (Subject subject in myStructure.GetSubjectList())
                            {
                                if (subject.GetsubjectID() == teacherSubject.GetSubjectID()
                                    && teacher.GetID() == teacherSubject.GetTeacherID()
                                    && subject.GetsubjectTitle() == teacherSubjectEditSubjectListBox.SelectedItem.ToString())
                                {
                                    Added = true;
                                    break;
                                }
                            }
                        }
                        break;
                    }
                }
                if (!Added)
                {
                    foreach (Teacher teacher in myStructure.GetTeacherList())
                    {
                        if (teacher.Getname() == checkedName)
                        {
                            foreach (Subject subject in myStructure.GetSubjectList())
                            {
                                if (subject.GetsubjectTitle() == teacherSubjectEditSubjectListBox.SelectedItem.ToString())
                                {
                                    myStructure.GetTeacherSubjectList().Add(new TeacherSubject(teacher.GetID(), subject.GetsubjectID()));
                                    break;
                                }
                            }
                            break;
                        }
                    }
                }
            }
        }

        private void teacherSubjectMainMenuButton_Click(object sender, EventArgs e)
        {
            MakeClass();
            MakeTimetable();
            teacherSubjectEditSubjectListBox.Visible = false;
            teacherSubjectEditTeacherCheckedListBox.Visible = false;
            teacherSubjectSaveSpecialismButton.Visible = false;
            TimetableLayoutPanel1.Visible = true;
            teacherSubjectMainMenuButton.Visible = false;
        }

        private void UserPreferencesButton_Click(object sender, EventArgs e)
        {
            TimetableLayoutPanel1.Visible = false;
            userPreferencesItemsPerDayLabel.Visible = true;
            userPreferencesItemsPerDayNumericUpDown.Value = myStructure.GetitemsPerDay();
            userPreferencesItemsPerDayNumericUpDown.Visible = true;
            userPreferencesSavePreferencesButton.Visible = true;
            userPreferencesCancelButton.Visible = true;
        }

        private void AddSubjectButton_Click(object sender, EventArgs e)
        {
            TimetableLayoutPanel1.Visible = false;
            subjectAddSubjectLabel.Visible = true;
            subjectAddSubjectComboBox.Items.Clear();
            foreach (Subject subject in myStructure.GetSubjectList())
            {
                subjectAddSubjectComboBox.Items.Add(subject.GetsubjectTitle());
            }
            subjectAddSubjectComboBox.Visible = true;
            subjectAddSubjectButton.Visible = true;
            subjectAddCancelButton.Visible = true;
        }

        private void subjectAddSubjectButton_Click(object sender, EventArgs e)
        {
            if (subjectAddSubjectComboBox.Text != "")
            {
                CreateSubject(subjectAddSubjectComboBox.Text);
                subjectAddSubjectLabel.Visible = false;
                subjectAddSubjectComboBox.Visible = false;
                subjectAddSubjectButton.Visible = false;
                subjectAddCancelButton.Visible = false;
                TimetableLayoutPanel1.Visible = true;
                MakeTimetable();
                GetExcelFile();
            }
            else
            {
                MessageBox.Show("Invalid Input");
            }
        }

        private void subjectAddCancelButton_Click(object sender, EventArgs e)
        {
            subjectAddCancelButton.Visible = false;
            subjectAddSubjectButton.Visible = false;
            subjectAddSubjectComboBox.Visible = false;
            subjectAddSubjectLabel.Visible = false;
            TimetableLayoutPanel1.Visible = true;
        }

        private void propertiesAddBlockButton_Click(object sender, EventArgs e)
        {
            if (propertiesPeriodTitleTextBox.Text != null
                && propertiesPeriodLengthNumericUpDown.Value != 0)
            {
                int highestID = 0;
                foreach (ScheduleBlock scheduleBlock in myStructure.GetScheduleBlockList())
                {
                    if (scheduleBlock.GetID() > highestID)
                    {
                        highestID = scheduleBlock.GetID();
                    }
                }
                myStructure.GetScheduleBlockList().Add(new ScheduleBlock(highestID + 1, propertiesPeriodTitleTextBox.Text, new TimeSpan(0, Convert.ToInt32(propertiesPeriodLengthNumericUpDown.Value),0), propertiesSchedulableCheckBox.Checked));
                propertiesDayStartTimedateTimePicker.Format = DateTimePickerFormat.Custom;
                // Display the time as "10:15".  
                propertiesDayStartTimedateTimePicker.CustomFormat = "HH:mm";
                myStructure.SetschoolStartTime(propertiesDayStartTimedateTimePicker.Value);
                propertiesTimetableTableLayoutPanel.Controls.Clear();
                propertiesTimetableTableLayoutPanel.ColumnCount = 0;
                propertiesTimetableTableLayoutPanel.ColumnStyles.Clear();
                TimeSpan timeCount = new TimeSpan();
                DateTime startTime = myStructure.GetschoolStartTime().GetValueOrDefault(DateTime.Now);
                foreach (ScheduleBlock scheduleBlock in myStructure.GetScheduleBlockList())
                {
                    scheduleBlock.SetstartTime(startTime.Add(timeCount));
                    Label label = new Label();
                    string Time = (startTime.Add(timeCount)).ToString("HH:mm") + " - " + (startTime.Add(timeCount)).Add(scheduleBlock.GetperiodLength()).ToString("HH:mm");
                    timeCount = timeCount.Add(scheduleBlock.GetperiodLength());
                    label.Text = scheduleBlock.GetperiodTitle() + "\n" +  Time;
                    label.AutoSize = true;
                    propertiesTimetableTableLayoutPanel.ColumnCount += 1;
                    propertiesTimetableTableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
                    propertiesTimetableTableLayoutPanel.Controls.Add(label);
                }
            }
            else
            {
                MessageBox.Show("Invalid Input");
            }
        }

        private void propertiesDeleteBlockButton_Click(object sender, EventArgs e)
        {
            if (myStructure.GetScheduleBlockList().Count != 0)
            {
                myStructure.GetScheduleBlockList().RemoveAt(myStructure.GetScheduleBlockList().Count - 1);
                propertiesTimetableTableLayoutPanel.Controls.Clear();
                propertiesTimetableTableLayoutPanel.ColumnCount = 0;
                propertiesTimetableTableLayoutPanel.ColumnStyles.Clear();
                TimeSpan timeCount = new TimeSpan();
                DateTime startTime = myStructure.GetschoolStartTime().GetValueOrDefault(DateTime.Now);
                foreach (ScheduleBlock scheduleBlock in myStructure.GetScheduleBlockList())
                {
                    Label label = new Label();
                    string Time = (startTime + timeCount).ToString("HH:mm") + " - " + (startTime + timeCount).Add(scheduleBlock.GetperiodLength()).ToString("HH:mm");
                    timeCount = timeCount.Add(scheduleBlock.GetperiodLength());
                    label.Text = scheduleBlock.GetperiodTitle() + "\n" + Time;
                    label.AutoSize = true;
                    propertiesTimetableTableLayoutPanel.ColumnCount += 1;
                    propertiesTimetableTableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
                    propertiesTimetableTableLayoutPanel.Controls.Add(label);
                }
            }
            else
            {
                MessageBox.Show("No Scheduled Blocks");
            }
        }

        private void userPreferencesSavePreferencesButton_Click(object sender, EventArgs e)
        {
            myStructure.SetitemsPerDay(Convert.ToInt32(userPreferencesItemsPerDayNumericUpDown.Value));
            foreach (User user in myStructure.GetUserList())
            {
                if (user.GetID() == globalCurrentUserID)
                {
                    user.SettasksPerDay(Convert.ToInt32(userPreferencesItemsPerDayNumericUpDown.Value));
                    break;
                }
            }
            userPreferencesItemsPerDayLabel.Visible = false;
            userPreferencesItemsPerDayNumericUpDown.Visible = false;
            userPreferencesCancelButton.Visible = false;
            userPreferencesSavePreferencesButton.Visible = false;
            TimetableLayoutPanel1.Visible = true;
            MakeTimetable();
        }

        private void userPreferencesCancelButton_Click(object sender, EventArgs e)
        {
            userPreferencesItemsPerDayLabel.Visible = false;
            userPreferencesItemsPerDayNumericUpDown.Visible = false;
            userPreferencesCancelButton.Visible = false;
            userPreferencesSavePreferencesButton.Visible = false;
            TimetableLayoutPanel1.Visible = true;
        }

    }
}
