namespace GptAssistant.Data.Models
{
    public class Message
    {
        public string Role { get; set; }
        public List<ContentBlock> Content { get; set; }
    }
}
