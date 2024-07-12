namespace GptAssistant.Service.AiService
{
    public interface IAiService
    {
        Task<string> AskGptQuestion(string question);
        Task ProcessExcelAndAskGpt(string excelFilePath);
        List<string> ReadExcelQuestions(string excelFilePath);
        void WriteAnswersToExcel(string excelFilePath, List<string> questions, List<string> answers);
    }
}
