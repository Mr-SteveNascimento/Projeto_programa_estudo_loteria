from xlwt import Workbook
import xlrd

def calculo():
	ultima_linha = ultima()
	matriz = arquivo_entrada()
	p = 0
	workbook = xlrd.open_workbook('resultados.xlsx')
	worksheet = workbook.sheet_by_index(0)
	linhas = worksheet.nrows
	colunas = worksheet.ncols
	for i in range(linhas-ultima_linha):
		w = i + ultima_linha
		p += 1
		if worksheet.cell(w,10).value != '0':
			valor = list(worksheet.cell(w,10).value)
			for t in range(len(valor)):
				valor[t] = int(valor[t])
			for premio in range(10):
				for dezena in range(10):
					vetor = [premio, dezena]
					if vetor[0] == valor[0] and vetor[1] == valor[1]:
						if matriz[premio][dezena][5] > 1:
							matriz[premio][dezena][6].append(matriz[premio][dezena][5])
						matriz[premio][dezena][5] = 0
						matriz[premio][dezena][0] += 1
						matriz[premio][dezena][4] += 1
						if matriz[premio][dezena][4] > matriz[premio][dezena][2]:
							matriz[premio][dezena][2] = matriz[premio][dezena][4]
					elif vetor[0] == valor[0] and vetor[1] != valor[1]:
						matriz[premio][dezena][4] = 0
						matriz[premio][dezena][1] += 1
						matriz[premio][dezena][5] += 1
						if matriz[premio][dezena][5] > matriz[premio][dezena][3]:
							matriz[premio][dezena][3] = matriz[premio][dezena][5]
		print(p, '%')
	ultima_pqp(linhas)
	return(matriz)
def planilha():
	workbook1 = Workbook()
	worksheet1 = workbook1.add_sheet('Planilha1')
	lista_de_contador_de_contador = []
	matrix = calculo()
	for i in range(10):
		for j in range(10):
			quantidade = matrix[i][j][6].count(matrix[i][j][5])
			lista_de_contador_de_contador.append(quantidade)
	controle = []
	controle2 = []
	for i in range(10):
		controle.append(lista_de_contador_de_contador[i])
		controle2.append(lista_de_contador_de_contador[-i-1])
		lista_de_contador_de_contador.pop()
		del(lista_de_contador_de_contador[0])
	controle.reverse()
	for i in range(10):
		lista_de_contador_de_contador.append(controle2[i])
		lista_de_contador_de_contador.insert(0, controle[i])
	worksheet1.write(0, 1, "1")
	worksheet1.write(0, 11, "2")
	worksheet1.write(0, 21, "3")
	worksheet1.write(0, 31, "4")
	worksheet1.write(0, 41, "5")
	worksheet1.write(0, 51, "6")
	worksheet1.write(0, 61, "7")
	worksheet1.write(0, 71, "8")
	worksheet1.write(0, 81, "9")
	worksheet1.write(0, 91, "10")

	for linha in range(6):
		linhatotal = linha + 2
		for colunatotal in range(92):
			if colunatotal == 1:
				for coluna in range(10):
					worksheet1.write(linhatotal, colunatotal+coluna, matrix[1][coluna][linha])
					lista = matrix[1][coluna][6]
			elif colunatotal == 11:
				for coluna in range(10):
					worksheet1.write(linhatotal, colunatotal+coluna, matrix[2][coluna][linha])
					lista = matrix[2][coluna][6]
			elif colunatotal == 21:
				for coluna in range(10):
					worksheet1.write(linhatotal, colunatotal+coluna, matrix[3][coluna][linha])
					lista = matrix[3][coluna][6]
			elif colunatotal == 31:
				for coluna in range(10):
					worksheet1.write(linhatotal, colunatotal+coluna, matrix[4][coluna][linha])
					lista = matrix[4][coluna][6]
			elif colunatotal == 41:
				for coluna in range(10):
					worksheet1.write(linhatotal, colunatotal+coluna, matrix[5][coluna][linha])
					lista = matrix[5][coluna][6]
			elif colunatotal == 51:
				for coluna in range(10):
					worksheet1.write(linhatotal, colunatotal+coluna, matrix[6][coluna][linha])
					lista = matrix[6][coluna][6]
			elif colunatotal == 61:
				for coluna in range(10):
					worksheet1.write(linhatotal, colunatotal+coluna, matrix[7][coluna][linha])
					lista = matrix[7][coluna][6]
			elif colunatotal == 71:
				for coluna in range(10):
					worksheet1.write(linhatotal, colunatotal+coluna, matrix[8][coluna][linha])
					lista = matrix[8][coluna][6]
			elif colunatotal == 81:
				for coluna in range(10):
					worksheet1.write(linhatotal, colunatotal+coluna, matrix[9][coluna][linha])
					lista = matrix[9][coluna][6]
			elif colunatotal == 91:
				for coluna in range(10):
					worksheet1.write(linhatotal, colunatotal+coluna, matrix[0][coluna][linha])
					lista = matrix[0][coluna][6]
	worksheet1.write(2, 0, "Quantidade Sim")
	worksheet1.write(3, 0, "Quantidade Não")
	worksheet1.write(4, 0, "Maior Sequencia de Sim")
	worksheet1.write(5, 0, "Maior Sequencia de Não")
	worksheet1.write(6, 0, "Valor atual de contador Sim")
	worksheet1.write(7, 0, "Valor atual de contador Não")
	for colunatotal in range(100):
		worksheet1.write(8, colunatotal+1, lista_de_contador_de_contador[colunatotal])
	for colunatotal in range(92):
		if colunatotal == 1:
			for i in range(10):
				lista = matrix[1][i][6]
				for j in range(len(lista)):
					worksheet1.write(9+j, colunatotal+i, lista[j])
		elif colunatotal == 11:
			for i in range(10):
				lista = matrix[2][i][6]
				for j in range(len(lista)):
					worksheet1.write(9+j, colunatotal+i, lista[j])
		elif colunatotal == 21:
			for i in range(10):
				lista = matrix[3][i][6]
				for j in range(len(lista)):
					worksheet1.write(9+j, colunatotal+i, lista[j])
		elif colunatotal == 31:
			for i in range(10):
				lista = matrix[4][i][6]
				for j in range(len(lista)):
					worksheet1.write(9+j, colunatotal+i, lista[j])
		elif colunatotal == 41:
			for i in range(10):
				lista = matrix[5][i][6]
				for j in range(len(lista)):
					worksheet1.write(9+j, colunatotal+i, lista[j])
		elif colunatotal == 51:
			for i in range(10):
				lista = matrix[6][i][6]
				for j in range(len(lista)):
					worksheet1.write(9+j, colunatotal+i, lista[j])
		elif colunatotal == 61:
			for i in range(10):
				lista = matrix[7][i][6]
				for j in range(len(lista)):
					worksheet1.write(9+j, colunatotal+i, lista[j])
		elif colunatotal == 71:
			for i in range(10):
				lista = matrix[8][i][6]
				for j in range(len(lista)):
					worksheet1.write(9+j, colunatotal+i, lista[j])
		elif colunatotal == 81:
			for i in range(10):
				lista = matrix[9][i][6]
				for j in range(len(lista)):
					worksheet1.write(9+j, colunatotal+i, lista[j])
		elif colunatotal == 91:
			for i in range(10):
				lista = matrix[0][i][6]
				for j in range(len(lista)):
					worksheet1.write(9+j, colunatotal+i, lista[j])
	workbook1.save("Saída.xls")
	arquivo_saida(matrix)
def arquivo_entrada():
	arquivoI = open('arquivoI.txt', 'r')
	arquivoII = open('arquivoII.txt', 'r')
	lista_arquivoI = []
	lista_arquivoII = []
	lista_saida = [[], [], [], [], [], [], [], [], [], []]
	lista_controle_saida = []
	for linha in arquivoI:
		valores = linha.split()
		valores.pop()
		lista_arquivoI.append(valores)
	for linha in arquivoII:
		valores = linha.split()
		valores.pop()
		lista_arquivoII.append(valores)
	for valor in range(len(lista_arquivoI)):
		lista_arquivoI[valor].append(lista_arquivoII[valor])
		for i in range(len(lista_arquivoI[valor][6])):
			lista_arquivoI[valor][6][i] = int(lista_arquivoI[valor][6][i])
		for i in range(6):
			lista_arquivoI[valor][i] = int(lista_arquivoI[valor][i])
	contador = 0
	casa = 0
	for i in range(len(lista_arquivoI)):
		lista_saida[casa].append(lista_arquivoI[i])
		contador += 1
		if contador == 10:
			casa += 1
			contador = 0
	arquivoI.close()
	arquivoII.close()
	return(lista_saida)
def arquivo_saida(matriz):
	arquivoI = open('arquivoI.txt', 'w')
	arquivoII = open('arquivoII.txt', 'w')
	matriz_direita = []
	lista_arquivoII = []
	lista_para_escrever_I = []
	lista_para_escrever_II = []
	for i in range(10): #Coloca todos os valores em ordem em uma lista total
		for j in range(10):
			matriz_direita.append(matriz[i][j])
	for i in range(100): #Separa a lista de atrasos dos outros valores
		lista_arquivoII.append(matriz_direita[i][6])
		matriz_direita[i].pop()
	for i in range(len(matriz_direita)): #Transforma a matriz em str e adiciona o valor a ser excluido
		matriz_direita[i].append('0\n')
		for j in range(6):
			matriz_direita[i][j] = str(matriz_direita[i][j])
	for i in range(len(lista_arquivoII)):
		lista_arquivoII[i].append('0\n')
		for j in range(len(lista_arquivoII[i])):
			lista_arquivoII[i][j] = str(lista_arquivoII[i][j])
	for i in range(len(matriz_direita)):
		string = ' '.join(matriz_direita[i])
		lista_para_escrever_I.append(string)
	for i in range(len(lista_arquivoII)):
		string = ' '.join(lista_arquivoII[i])
		lista_para_escrever_II.append(string)
	arquivoI.writelines(lista_para_escrever_I)
	arquivoII.writelines(lista_para_escrever_II)
	arquivoI.close()
	arquivoII.close()
def ultima():
	arquivo = open('linha.txt', 'r')
	linha = arquivo.readline()
	linha = int(linha)
	print(linha)
	arquivo.close()
	return(linha)
def ultima_pqp(pqp):
	linha = str(pqp)
	arquivo = open('linha.txt', 'w')
	p = [linha]
	arquivo.writelines(p)
	arquivo.close()
planilha()
