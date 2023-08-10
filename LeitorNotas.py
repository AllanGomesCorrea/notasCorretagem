# This is a code created to generate Excel sheet with all main data in NotasCorretagem.pdf
# The main data is located at the bottom of PDF
# Notas de Corretagem used from MODALMAIS Corretora ltda.
# Version 1.0 created by Allan Gomes Corrêa

if __name__ == '__main__':

    import tabula
    import pandas as pd
    import PyPDF2

    #PDF path
    pdf_path = 'NotaCorretagem2.pdf'

    #Count number of pages
    with open(pdf_path, 'rb') as pdf_file:
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        num_pages = len(pdf_reader.pages)

    # Create dataframe based on NotaCorretagem fields.
    data = ['Data', 'Venda disponível', 'Compra disponível', 'Venda opções', 'Compra opções', 'Valor dos negócios',
            'IRRF', 'IRRF Day Trade(Projeção)', 'Taxa operacional', 'Taxa registro BMF', 'Taxa BMF (emol+f.gar)',
            'Outros Custos', 'ISS', 'Ajuste de posição', 'Ajuste Day Trade', 'Total das despesas',
            'Outros', 'IRRF Corretagem', 'Total Conta Investimento', 'Total Conta Normal', 'Total líquido',
            'Total líquido da nota']
    df = pd.DataFrame(data)
    df = df.transpose()

    for page in range(1,num_pages+1):
        # Extract tables from PDF
        tables = tabula.read_pdf(pdf_path, pages=page, multiple_tables=True)
        first_table = tables[0]
        last_table = tables[len(tables)-1]

        data_extracted = [first_table.iloc[1, 4]]

        #Add all data extracted from last table desired
        for index, row in last_table.iterrows():
            for column, value in row.items():
                if isinstance(value,str) and value.find(",") != -1:
                    data_extracted.append(value)


        #Convert str numbers to float and apply Credit or Debit for each field
        for i in range(1,len(data_extracted)):
            aux = data_extracted[i].find("|")
            if aux == -1:
                data_extracted[i] = float(data_extracted[i].replace(',', '.'))
            else:
                num = data_extracted[i][:data_extracted[i].find("|")-1]
                num = float(num.replace(',', '.'))
                inum = data_extracted[i][data_extracted[i].find("|")+2:]
                if inum == 'C':
                    data_extracted[i] = num * 1
                elif inum == 'D':
                    data_extracted[i] = num*-1
                else:
                    data_extracted[i] = num
        # Add to dataframe
        df.loc[page] = data_extracted

    #Create .csv with all data
    arq_name = 'Planilha_NotasCorretagem.xlsx'
    df.to_excel(arq_name, index=False)
