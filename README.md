# 📊 SelectQueryToExcelHelper

![GitHub release](https://img.shields.io/badge/version-1.0-blue)
![License](https://img.shields.io/badge/license-MIT-green)

**SelectQueryToExcelHelper** is a simple **C# command-line tool** that reads SQL queries from a file, executes them, and exports the results into an **Excel file (`.xlsx`)**. Each query is stored in a separate sheet with headers, and the executed SQL query is added at the bottom.

---

## 🚀 Features
✅ Read SQL queries from a `.sql` or `.txt` file  
✅ Execute queries against a **SQL Server database**  
✅ Save results in an **Excel file**, with **each query in a separate sheet**  
✅ Include **column headers**, even if no data is returned  
✅ Append the **executed SQL query** at the bottom of the sheet  
✅ **Self-contained EXE** – No .NET installation required  

---

## 📥 Download
🔗 [**Download the latest version**](https://drive.google.com/drive/folders/1iLe_V9mvS-9yOYg6Ix9TOwSbXN4nozUN?usp=sharing)  
*(No installation required – just download and run!)*

---

## 🛠 How to Use

### **1️⃣ Setup**
- Ensure you have a `.sql` or `.txt` file containing **SELECT queries**, separated by `;`.
- Create a `db_connection.txt` file with your **SQL Server connection string**.

### **2️⃣ Running the Application**
- **Double-click** `SelectQueryToExcelHelper.exe` and follow the prompts.  
- **Or run via Command Prompt**:
  ```sh
  SelectQueryToExcelHelper.exe

### **3️⃣ Example Usage**

#### **📂 db_connection.txt**
Create a `db_connection.txt` file in the same directory as the EXE. It should contain your SQL Server connection string:
```txt
Server=myServerAddress;Database=myDataBase;User Id=myUsername;Password=myPassword;
```

#### **📂 queries.sql**
Write your SQL queries in a `.sql` or `.txt` file, separating multiple queries with `;`:

```sql
SELECT TOP 10 * FROM Employees;
SELECT * FROM Orders WHERE OrderID > 100;
```
#### **📂 Output: queries.xlsx**
Each query is saved in a **separate sheet** in the Excel file.

| **Sheet Name** | **Content** |
|---------------|-------------|
| `Sheet1`      | First 10 employees                  |
| `Sheet2`      | Orders where `OrderID > 100`        |

Each sheet includes **headers** and the **executed query** at the bottom.

---

## 🏗 Building from Source

### **🔧 Prerequisites**
- .NET 6 SDK or higher
- Visual Studio / VS Code

### **📦 Clone & Build**
Run the following commands in your terminal or command prompt:

```sh
git clone https://github.com/YOUR_GITHUB_USERNAME/SelectQueryToExcelHelper.git
cd SelectQueryToExcelHelper
dotnet publish -c Release -r win-x64 --self-contained true
```
This will generate a standalone executable in the bin/Release/net6.0/win-x64/publish folder.

## 📜 License
This project is licensed under the MIT License – see the [LICENSE](https://choosealicense.com/licenses/mit/) file for details.

## 💡 Contributions & Feedback
Contributions are welcome! If you have suggestions or feature requests:

Open an issue on GitHub
Submit a pull request
Share your feedback!
