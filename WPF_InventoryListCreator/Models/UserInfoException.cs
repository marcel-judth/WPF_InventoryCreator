using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WPF_InventoryListCreator.Models
{
    class UserInfoException : Exception
    {
        public UserInfoException()
        {
        }

        public UserInfoException(string message)
            : base(message)
        {
        }
    }
}
