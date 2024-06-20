"""
pegar dados da planilha 
Nome do Curso	Nome Participante	Tipo de Participação	Data de Início	Data de Término	Carga Horária (horas)	Data de Emissão do Certificado

trasnferir para a imagem do certificado

"""

import openpyxl
from PIL import Image, ImageDraw, ImageFont 

#abrir a planilha
workbook_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')
sheet_alunos = workbook_alunos['Sheet1'] 

for indice, linha in enumerate(sheet_alunos.iter_rows(min_row=2)) :
    nome_curso = linha[0].value #nome do curso
    nome_participante = linha[1].value #nome participante
    tipo_participacao = linha[2].value 
    data_inicio = linha[3].value 
    data_fim = linha[4].value 
    carga_horaria = linha[5].value 
    data_emissao_certificado = linha[6 ].value
    # transferir os dados da planilha para a imagem do certificado

    #definindo fonte a ser utilizada
    font_nome = ImageFont.truetype('./tahomabd.ttf', 90)
    font_geral = ImageFont.truetype('./tahoma.ttf', 80)
    font_data = ImageFont.truetype('./tahoma.ttf', 55)

    image = Image.open('./certificado_padrao.jpg')
    desenhar = ImageDraw.Draw(image)

    desenhar.text((1010, 829), nome_participante, fill='black', font=font_nome)
    desenhar.text((1060, 950), nome_curso, fill='black', font=font_geral)
    desenhar.text((1435, 1065), tipo_participacao, fill='black', font=font_geral)
    desenhar.text((750, 1770), str(data_inicio), fill='blue', font=font_data)
    desenhar.text((750, 1930), str(data_fim), fill='blue', font=font_data)

    desenhar.text((2220, 1930), str(data_emissao_certificado), fill='blue', font=font_data)



    image.save(f'./certificados/{indice} {nome_participante} - certificado.png')



