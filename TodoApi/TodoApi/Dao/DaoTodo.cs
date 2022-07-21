using Microsoft.Data.SqlClient;
using System.Data;
using TodoApi.Models;

namespace TodoApi.Dao

{
    public class DaoTodo
    {
        string conexao = @"Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=Todos;Integrated Security=True";
        public List<TodoItem> GetTodoItems()
        {
            List<TodoItem> todoitem = new List<TodoItem>();
            using (SqlConnection conn = new SqlConnection(conexao))
            {
                
                using (SqlCommand cmd = new SqlCommand("SELECT * FROM TodoItems",conn))
                {
                    conn.Open();
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if(reader != null)
                        {
                            while (reader.Read())
                            {
                                Console.Write(reader["IsComplete"].ToString() == IsComplete.True.ToString());
                                var todoitems = new TodoItem();
                                todoitems.Name = reader["Name"].ToString();
                                todoitems.IsComplete = reader["IsComplete"].ToString();
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
