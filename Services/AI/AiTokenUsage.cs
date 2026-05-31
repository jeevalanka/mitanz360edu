namespace MITANZ360Edu.Web.Services.AI;

public sealed class AiTokenUsage
{
    public int PromptTokens { get; set; }

    public int CompletionTokens { get; set; }

    public int TotalTokens { get; set; }
}