using Microsoft.Data.SqlClient;
using System.Data;
using TodoApi.Models;

namespace TodoApi.Dao

{
    public class DaoTodo
    {
        string conexao = @"Data Source=DESKTOP-HV9333S;Initial Catalog=Todos;Integrated Security=True";
        public List<TodoItem> GetTodoItems()
        {
            List<TodoItem> todoitem = new List<TodoItem>();
            using (SqlConnection conn = new SqlConnection(conexao))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("SELECT * FROM TodoItems",conn))
                {
                    cmd.CommandType = CommandType.Text;
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if(reader != null)
                        {
                            while (reader.Read())
                            {
                                var todoitems = new TodoItem();
                                todoitems.Name = reader["Nome"].ToString();
                                todoitems.IsComplete = reader["IsComplete"].is;
                                todoitem.Add(todoitems);
                               
                            }
                        }
                    }
                }
            }
                return todoitem;
        }
    }
}
