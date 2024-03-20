using System;
using System.Linq;
using ClosedXML.Excel;

namespace ConsoleApp
{
	class Program
	{
		static void Main(string[] args)
		{
			string filePath = string.Empty;
			XLWorkbook workbook = null;

			// Запрос на ввод пути до файла с данными
			Console.WriteLine("Введите путь до файла с данными:");
			filePath = Console.ReadLine();

			try
			{
				workbook = new XLWorkbook(filePath);
			}
			catch (Exception ex)
			{
				Console.WriteLine($"Ошибка при открытии файла: {ex.Message}");
				return;
			}

			// Основной цикл программы
			while (true)
			{
				Console.WriteLine("Выберите действие:");
				Console.WriteLine("1. Поиск информации о клиентах по наименованию товара");
				Console.WriteLine("2. Изменение контактного лица клиента");
				Console.WriteLine("3. Определение золотого клиента");
				Console.WriteLine("4. Выход");

				string choice = Console.ReadLine();

				switch (choice)
				{
					case "1":
						SearchByProductName(workbook);
						break;
					case "2":
						ChangeContactPerson(workbook);
						break;
					case "3":
						DetermineGoldenCustomer(workbook);
						break;
					case "4":
						Console.WriteLine("Программа завершена.");
						return;
					default:
						Console.WriteLine("Некорректный ввод. Повторите попытку.");
						break;
				}
			}
		}

		static void SearchByProductName(XLWorkbook workbook)
		{
			Console.WriteLine("Введите наименование товара:");
			string productName = Console.ReadLine();

			var ordersWorksheet = workbook.Worksheet("Заявки");
			var productsWorksheet = workbook.Worksheet("Товары");
			var customersWorksheet = workbook.Worksheet("Клиенты");

			var productCode = productsWorksheet
							.Rows()
							.FirstOrDefault(r => r.Cell(2).Value.ToString().Equals(productName))
							?.Cell(1).Value.ToString();

			var orders = ordersWorksheet.RowsUsed()
				.Where(r => r.Cell(2).Value.ToString() == productCode)
				.Select(r => new
				{
					CustomerCode = r.Cell(3).Value.ToString(),
					Quantity = int.Parse(r.Cell(5).Value.ToString()),
					OrderDate = DateTime.Parse(r.Cell(6).Value.ToString())
				});

			Console.WriteLine($"Информация о клиентах, заказавших товар '{productName}':");
			foreach (var order in orders)
			{
				var customer = customersWorksheet.RowsUsed()
					.FirstOrDefault(r => r.Cell(1).Value.ToString() == order.CustomerCode);

				Console.WriteLine($"Клиент: {customer.Cell(2).Value}");
				Console.WriteLine($"Количество товара: {order.Quantity}");
				Console.WriteLine($"Дата заказа: {order.OrderDate.ToShortDateString()}\n");
			}
		}

		static void ChangeContactPerson(XLWorkbook workbook)
		{
			Console.WriteLine("Введите название организации:");
			string companyName = Console.ReadLine();
			Console.WriteLine("Введите ФИО нового контактного лица:");
			string newContactPerson = Console.ReadLine();

			var customersWorksheet = workbook.Worksheet("Клиенты");

			var customer = customersWorksheet.RowsUsed()
				.FirstOrDefault(r => r.Cell(2).Value.ToString() == companyName);

			if (customer != null)
			{
				customer.Cell(4).Value = newContactPerson;
				workbook.Save();
				Console.WriteLine("Информация о контактном лице изменена.");
			}
			else
			{
				Console.WriteLine("Клиент с таким названием организации не найден.");
			}
		}

		static void DetermineGoldenCustomer(XLWorkbook workbook)
		{
			Console.WriteLine("Введите год:");
			int year = int.Parse(Console.ReadLine());
			Console.WriteLine("Введите месяц:");
			int month = int.Parse(Console.ReadLine());

			var ordersWorksheet = workbook.Worksheet("Заявки");
			var customersWorksheet = workbook.Worksheet("Клиенты");

			var orders = ordersWorksheet.RowsUsed().Skip(1)
				.Where(r => DateTime.Parse(r.Cell(6).Value.ToString()).Year == year &&
							DateTime.Parse(r.Cell(6).Value.ToString()).Month == month)
				.GroupBy(r => r.Cell(3).Value.ToString())
				.Select(g => new
				{
					CustomerCode = g.Key,
					OrdersCount = g.Count()
				});

			var goldenCustomer = orders.OrderByDescending(o => o.OrdersCount).FirstOrDefault();

			if (goldenCustomer != null)
			{
				var customer = customersWorksheet.RowsUsed()
					.FirstOrDefault(r => r.Cell(1).Value.ToString() == goldenCustomer.CustomerCode);

				if (customer != null)
				{
					Console.WriteLine($"Золотой клиент: {customer.Cell(2).Value}");
					Console.WriteLine($"Количество заказов за {month}.{year}: {goldenCustomer.OrdersCount}");
				}
				else
				{
					Console.WriteLine("Информация о золотом клиенте не найдена.");
				}
			}
			else
			{
				Console.WriteLine("Заказов за указанный период не найдено.");
			}
		}
	}
}
