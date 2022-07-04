using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1
{
    public class Model
    {
        public string Name { get; set; }
        public int Quantity { get; set; }
    }

    public class Repo
    {
        public List<Model> RepoIntal()
        {
            var repo = new List<Model>() { new Model { Name="dsds", Quantity = 1 },
            new Model { Name="dsds", Quantity = 1 },
            new Model { Name="dsds", Quantity = 1 }};
            return repo;
        }
    }
}
