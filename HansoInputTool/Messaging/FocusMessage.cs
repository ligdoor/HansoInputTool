namespace HansoInputTool.Messaging
{
    // Viewにフォーカスを要求するためのメッセージ
    public class FocusMessage : IMessage
    {
        public string TargetElementName { get; set; }
    }
}