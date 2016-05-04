
#C# DataGridView 导出到Excel 的Demo 
**1.添加dll引用**

右击选择你所在的项目的“引用”，选择“添加引用”。

弹出“添加引用”对话框。

选择“COM”选项卡。

选择“Microsoft Excel 15.0 Object Library”（确保你的电脑里安装了Microsoft office 2007以上的版本）

单击“确定”按钮。

**2.代码部分**

首先“usingExcel = Microsoft.Office.Interop.Excel;”。

代码：
```

private void button1_Click(object sender, EventArgs e)    //button点击事件来调用
{    
if (ExportDataGridview(dataGridView1, true))    
MessageBox.Show("导出成功，请记得保存!");    
else   
MessageBox.Show("导出未成功，请检查是否有错!");    
}    
public bool ExportDataGridview(DataGridView gridView, bool isShowExcle)//生成Excel    
{    
if (gridView.Rows.Count == 0)    
return false;    
//建立Excel对象    
Excel.Application excel = new Excel.Application();    
excel.Application.Workbooks.Add(true);    
excel.Visible = isShowExcle;    
//生成字段名称    
for (int i = 0; i < gridView.ColumnCount; i++)    
{    
excel.Cells[1, i + 1] = gridView.Columns[i].HeaderText;    
}    
//填充数据    
for (int i = 0; i < gridView.RowCount; i++)    
{    
for (int j = 0; j < gridView.ColumnCount; j++)    
{    
if (gridView[j, i].ValueType == typeof(string))    
{    
excel.Cells[i + 2, j + 1] = "'" + gridView[j, i].Value.ToString();    
}    
else   
{    
excel.Cells[i + 2, j + 1] = gridView[j, i].Value.ToString();    
}    
}    
}    
return true;    
}
```
