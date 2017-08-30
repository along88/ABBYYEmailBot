using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ABBYYEmailBot
{
    public class EmailSender
    {

        public string name;
        public string ctrlNumber;
        public SenderStatus status;
    }

    public enum SenderStatus
    {
        NONE,
        SUCCESS,
        FAIL,
        PROCESSING
    }
}
