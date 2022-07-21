using TodoApi.Models;
using TodoApi.Dao;

namespace TodoApi.Repository
{
    public class TodoItemRepository
    {
        private readonly DaoTodo _daoTodoItem;
        public TodoItemRepository()
        {
            _daoTodoItem = new DaoTodo();
        }

        public List<TodoItem> GetTodoItems
        {
            get
            {
                return _daoTodoItem.GetTodoItems();
            }
        }
    }
}
