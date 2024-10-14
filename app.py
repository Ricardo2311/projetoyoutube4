# Enviar dados de uma planilha para campos de um certificado para cada aluno
# 1- Abrir uma planilha excel
# 2- Copiar cada célula da planilha
# 3- Abrir o certificado
# 4- Colar as informações nos campos certos dos certificados
# 5- Salvar os certificados com o nome de cada aluno
import openpyxl
from PIL import Image, ImageDraw, ImageFont


planilha = openpyxl.load_workbook('planilha_alunos.xlsx')
pagina = planilha['Sheet1']

for linha in pagina.iter_rows(min_row=2,values_only=True):
    nome_curso = linha[0]
    nome_participante = linha[1]
    tipo_participacao = linha[2]
    data_inicio = linha[3]
    data_fim = linha[4]
    carga_horaria = str(linha[5])
    data_emissao = linha[6]
    imagem = Image.open('certificado_padrao.jpg')
    draw = ImageDraw.Draw(imagem)
    posicao_nome_participante = (1010,843)
    posicao_nome_curso = (1065,963)
    posicao_tipo_participacao = (1430,1078)
    posicao_carga_horaria = (1490,1196)
    posicao_data_inicio = (700,1770)
    posicao_data_fim = (700,1920)
    posicao_data_emissao = (2190,1920)  
    fonte = ImageFont.truetype("arial.ttf",70)
    fonte_nome_participante = ImageFont.truetype("arialbd.ttf",70)
    draw.text(posicao_nome_curso,nome_curso,font=fonte,fill="black")
    draw.text(posicao_nome_participante,nome_participante,font=fonte_nome_participante,fill="black")
    draw.text(posicao_tipo_participacao,tipo_participacao,font=fonte,fill="black")
    draw.text(posicao_data_inicio,data_inicio,font=fonte,fill="black")
    draw.text(posicao_data_fim,data_fim,font=fonte,fill="black")
    draw.text(posicao_data_emissao,data_emissao,font=fonte,fill="black")
    draw.text(posicao_carga_horaria,carga_horaria,font=fonte,fill="black")
    imagem.save(f'./certificados_prontos/certificado_{nome_participante}.jpg')
