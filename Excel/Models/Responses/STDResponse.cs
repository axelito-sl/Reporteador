namespace Excel.Models.Responses
{
    public class STDResponse
    {
        public bool isError { get; set; }
        public string Message { get; set; }
        public int totalRecords { get; set; }
        public List<object> data { get; set; }
    }
}
