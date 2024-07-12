using ExcelDataReader;
using GptAssistant.Data.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using RestSharp;
using System.Data;

namespace GptAssistant.Service.AiService
{
    public class AiService : IAiService
    {
        private readonly string apiKey;
        private readonly string assistantId;

        public AiService()
        {
            apiKey = "YOUR_API_KEY";
            assistantId = "YOUR_ASSISTANT_ID";
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public async Task ProcessExcelAndAskGpt(string excelFilePath)
        {
            var questions = ReadExcelQuestions(excelFilePath);
            var answers = new List<string>();

            foreach (var question in questions)
            {
                string answer = await AskGptQuestion(question);
                answers.Add(answer);
            }

            WriteAnswersToExcel(excelFilePath, questions, answers);
        }

        public List<string> ReadExcelQuestions(string excelFilePath)
        {
            try
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                var questions = new List<string>();

                using (var stream = File.Open(excelFilePath, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                            {
                                UseHeaderRow = true
                            }
                        });

                        var table = result.Tables[0];

                        foreach (DataRow row in table.Rows)
                        {
                            questions.Add(row[0].ToString());
                        }
                    }
                }

                return questions;

            }
            catch (Exception)
            {

                throw;
            }
        }
        public void WriteAnswersToExcel(string excelFilePath, List<string> questions, List<string> answers)
        {
            try
            {
                var fileInfo = new FileInfo(excelFilePath);

                using (var package = new ExcelPackage(fileInfo))
                {
                    var worksheet = package.Workbook.Worksheets[0];

                    for (int i = 0; i < questions.Count; i++)
                    {
                        worksheet.Cells[i + 2, 1].Value = questions[i];
                        worksheet.Cells[i + 2, 2].Value = answers[i];
                    }

                    package.Save();
                }
            }
            catch (Exception)
            {

                throw;
            }
        }

        public async Task<string> AskGptQuestion(string question)
        {
            var client = new RestClient();
            var threadId = "YOUR_THREAD_ID";

            var messageResponse = await SendMessageAsync(client, threadId, question);
            if (!messageResponse.IsSuccessful)
            {
                return messageResponse.ErrorMessage;
            }

            var runId = await CreateRunAsync(client, threadId);
            if (runId == null)
            {
                return "Run not created.";
            }

            var runCompleted = await WaitForRunCompletionAsync(client, threadId, runId);
            if (!runCompleted)
            {
                return "Run completed.";
            }

            var answer = await GetAssistantResponseAsync(client, threadId);
            return answer ?? "No response from assistant";
        }

        private async Task<RestResponse> SendMessageAsync(RestClient client, string threadId, string question)
        {
            string messageEndpoint = $"https://api.openai.com/v1/threads/{threadId}/messages";

            var messageRequest = CreateRestRequest(messageEndpoint, Method.Post, apiKey);

            var messageData = new
            {
                role = "user",
                content = new[]
                {
                    new
                    {
                        type = "text",
                        text = question
                    }
                }
            };
            messageRequest.AddJsonBody(JsonConvert.SerializeObject(messageData));
            return await client.ExecuteAsync(messageRequest);
        }

        private async Task<string> CreateRunAsync(RestClient client, string threadId)
        {
            string runEndpoint = $"https://api.openai.com/v1/threads/{threadId}/runs";

            var runRequest = CreateRestRequest(runEndpoint, Method.Post, apiKey);

            var runData = new
            {
                assistant_id = assistantId,
                additional_instructions = (string)null,
                tool_choice = (string)null
            };

            runRequest.AddJsonBody(JsonConvert.SerializeObject(runData));

            var runResponse = await client.ExecuteAsync(runRequest);
            if (!runResponse.IsSuccessful)
            {
                return null;
            }

            var runResult = JsonConvert.DeserializeObject<GptResponse>(runResponse.Content);
            return runResult?.Id;
        }

        private async Task<bool> WaitForRunCompletionAsync(RestClient client, string threadId, string runId)
        {
            string runStatus = "queued";
            while (runStatus != "completed")
            {
                await Task.Delay(2000);
                string runsEndpoint = $"https://api.openai.com/v1/threads/{threadId}/runs/{runId}";

                var runsRequest = CreateRestRequest(runsEndpoint, Method.Get, apiKey);

                var runsResponse = await client.ExecuteAsync(runsRequest);
                if (!runsResponse.IsSuccessful)
                {
                    return false;
                }

                var responseContent = JObject.Parse(runsResponse.Content);
                runStatus = responseContent["status"].ToString();
            }
            return true;
        }

        private async Task<string> GetAssistantResponseAsync(RestClient client, string threadId)
        {
            string messagesEndpoint = $"https://api.openai.com/v1/threads/{threadId}/messages";
            var messagesRequest = CreateRestRequest(messagesEndpoint, Method.Get, apiKey);

            var messagesResponse = await client.ExecuteAsync<ThreadMessagesResponse>(messagesRequest);
            if (messagesResponse.IsSuccessful)
            {
                var messages = messagesResponse.Data;
                foreach (var message in messages.Data)
                {
                    if (message.Role == "assistant" && message.Content != null)
                    {
                        foreach (var contentBlock in message.Content)
                        {
                            if (contentBlock.Type == "text")
                            {
                                return contentBlock.Text.Value;
                            }
                        }
                    }
                }
            }
            return null;
        }

        private RestRequest CreateRestRequest(string endpoint, Method method, string apiKey)
        {
            var request = new RestRequest(endpoint, method);
            request.AddHeader("Authorization", $"Bearer {apiKey}");
            request.AddHeader("Content-Type", "application/json");
            request.AddHeader("OpenAI-Beta", "assistants=v1");

            return request;
        }
    }
}
