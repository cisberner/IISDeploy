using System;
using System.Collections.Generic;

namespace IISDeploy.Core
{
    public class DeploymentResult
    {
        public bool Success { get; set; }
        public string Message { get; set; } = string.Empty;
        public List<string> LogMessages { get; private set; } = new List<string>();
        public Exception? Exception { get; set; }

        public DeploymentResult(bool success = false, string message = "")
        {
            Success = success;
            Message = message;
        }

        public void AddLog(string logMessage)
        {
            LogMessages.Add(logMessage);
        }
    }
}
