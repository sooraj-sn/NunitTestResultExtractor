namespace TestResultExtractor
{
    public class TestResults
    {
      
        public int testCount
        {
            get;
            set;
        }
        public string testCaseName
        {
            get;
            set;
        }

        public string executed
        {
            get;
            set;
        }

        public string result
        {
            get;
            set;
        }

        public string time
        {
            get;
            set;
        }
        public string failureReason
        {
            get;
            set;
        }

    }
}
