import pandas as pd
import xlrd
from pandas import read_excel

class Plastek:
    def __init__(self):
        pass
    
    def execute(self, path, output, excecao):
        self.path = path
        self.output = output
        self.excecao = excecao

        print('Path input:', self.path)
        print('Path output:', self.output)
        print('Abrindo o arquivo xlsx')
        print('Criando o data frame...')

        xlsx = pd.read_excel(self.path) # abre o arquivo xlsx
        frame = pd.DataFrame(xlsx) # cria um dataframe geral do arquivo

        print('Data frame criado!')
        print('Abrindo arquivo principal de saída...')

        arq_all = open (str(self.output)+'Escala-geral.txt', 'a')

        print('Arquivo principal de saída criado!')
        print('Coletando dados...')

        t1e = frame.loc[0, 'GRUPO1 E']
        t1s = frame.loc[0, 'GRUPO1 S']
        t2e = frame.loc[0, 'GRUPO2 E']
        t2s = frame.loc[0, 'GRUPO2 S']
        t3e = frame.loc[0, 'GRUPO3 E']
        t3s = frame.loc[0, 'GRUPO3 S']
        t4e = frame.loc[0, 'GRUPO4 E']
        t4s = frame.loc[0, 'GRUPO4 S']

        ida = frame.loc[0, 'ID A'].astype(int)
        idb = frame.loc[0, 'ID B'].astype(int)
        idc = frame.loc[0, 'ID C'].astype(int)

        ano = frame.loc[0, 'ANO'].astype(int)
        mes = frame.loc[0, 'MÊS'].astype(int)
        print('Dados coletados!')
        print('Iniciando os filtros e as escritas nos arquivos de saída...')

        dia = 0
        for i in frame['GRUPO1'].dropna():
            print('Entrou no for grupo1')
            dia=dia+1
            if (i == 1):
                arq_all.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(self.excecao)+"',"+str(t1e)+"','Exceção',"+str(t1s)+","+str(ida)+");"+'\n')
                
                arq = open (str(self.output)+'Grupo 1 mes-'+str(mes)+'.txt', 'a')
                arq.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(self.excecao)+"',"+str(t1e)+"','Exceção',"+str(t1s)+","+str(ida)+");"+'\n')
                arq.close ()
            
            
            
            elif (i == 0):
                arq_all.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(self.excecao)+"','null','Exceção','null',"+str(ida)+");"+'\n')
                
                arq = open (str(self.output)+'Grupo 1 mes-'+str(mes)+'.txt', 'a')
                arq.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(self.excecao)+"','null','Exceção','null',"+str(ida)+");"+'\n')
                arq.close ()
        print('Arquivo criado com sucesso!!')

        dia=0
        for i in frame['GRUPO2'].dropna():
            print('Entrou no for grupo2')    
            dia=dia+1
            if (i == 1):
                arq_all.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(self.excecao)+"',"+str(t2e)+",'Exceção',"+str(t2s)+","+str(idb)+");"+'\n')

                arq = open (str(self.output)+'Grupo 2 mes-'+str(mes)+'.txt', 'a')
                arq.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(self.excecao)+"',"+str(t2e)+",'Exceção',"+str(t2s)+","+str(idb)+");"+'\n')
                arq.close ()
            
            
            
            elif (i == 0):
                arq_all.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(self.excecao)+"','null','Exceção','null',"+str(ida)+");"+'\n')
                
                arq = open (str(self.output)+'Grupo 2 mes-'+str(mes)+'.txt', 'a')
                arq.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(self.excecao)+"','null','Exceção','null',"+str(ida)+");"+'\n')
                arq.close ()
        print('Arquivo criado com sucesso!!')

        dia=0
        for i in frame['GRUPO3'].dropna():
            print('Entrou no for grupo3')    
            dia=dia+1
            if (i == 1):
                arq_all.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(self.excecao)+"',"+str(t3e)+",'Exceção',"+str(t3s)+","+str(idc)+");"+'\n')

                arq = open (str(self.output)+'Grupo 3 mes-'+str(mes)+'.txt', 'a')
                arq.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(self.excecao)+"',"+str(t3e)+",'Exceção',"+str(t3s)+","+str(idc)+");"+'\n')
                arq.close ()
            
        
            
            elif (i == 0):
                arq_all.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(self.excecao)+"','null','Exceção','null',"+str(ida)+");"+'\n')
                
                arq = open (str(self.output)+'Grupo 3 mes-'+str(mes)+'.txt', 'a')
                arq.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(self.excecao)+"','null','Exceção','null',"+str(ida)+");"+'\n')
                arq.close ()
        print('Arquivo criado com sucesso!!')

        dia=0
        for i in frame['GRUPO4'].dropna():
            print('Entrou no for grupo4')    
            dia=dia+1
            if (i == 1):
                arq_all.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(self.excecao)+"',"+str(t4e)+",'Exceção',"+str(t4s)+","+str(ida)+");"+'\n')

                arq = open (str(self.output)+'Grupo 4 mes-'+str(mes)+'.txt', 'a')
                arq.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(self.excecao)+"',"+str(t4e)+",'Exceção',"+str(t4s)+","+str(ida)+");"+'\n')
                arq.close ()
            
            
            
            elif (i == 0):
                arq_all.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(self.excecao)+"','null','Exceção','null',"+str(ida)+");"+'\n')
                
                arq = open (str(self.output)+'Grupo 4 mes-'+str(mes)+'.txt', 'a')
                arq.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(self.excecao)+"','null','Exceção','null',"+str(ida)+");"+'\n')
                arq.close ()
        print('Arquivo criado com sucesso!!')

            
        arq_all.close()
        print('Arquivo criado com sucesso!!')
        print('Concluído!')

        return True