using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Example
{
    class User
    {
        public int? UserId { get; set; }
        public string Name { get; set; }
        public string Gender { get; set; }
        public DateTime? Birthday { get; set; }

        public User(int? UserId, string Name, string Gender, DateTime? Birthday)
        {
            this.UserId = UserId;
            this.Name = Name;
            this.Gender = Gender;
            this.Birthday = Birthday;
        }

        public User()
        {
        }
    }
}
