using Interfaces;
using System;

namespace HowToUse.Models
{
    class TestListEntry : IBaseModel
    {
        public int? ID { get; set; }
        public string Title { get; set; }
        public string Name { get; set; }
        public DateTime Erstellt_Am { get; set; }
        public int? ID_Database { get; set; }

        public override string ToString()
        {
            return $"{ID} {Title} {Name} {Erstellt_Am} {ID_Database}";
        }

    }
}
