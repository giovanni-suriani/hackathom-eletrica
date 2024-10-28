import win32com.client

# Caminho para o arquivo XLSM
caminho_arquivo = 'caminho_do_arquivo.xlsm'

# Iniciar o Excel
excel = win32com.client.Dispatch("Excel.Application")

# Abrir o arquivo XLSM
workbook = excel.Workbooks.Open(caminho_arquivo)

# Percorrer os módulos de VBA no projeto
for module in workbook.VBProject.VBComponents:
    if module.Type == 1:  # Tipo 1 é módulo padrão
        print(f'Módulo: {module.Name}')
        print(module.CodeModule.Lines(1, module.CodeModule.CountOfLines))

# Fechar o workbook sem salvar
workbook.Close(SaveChanges=False)

# Fechar o Excel
excel.Quit()
