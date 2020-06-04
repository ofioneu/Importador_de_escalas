import pandas as pd
import xlrd
from pandas import read_excel



class Viscofan:
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

        arq_all = open (str(self.output)+'escala.txt', 'a')

        print('Arquivo principal de saída criado!')
        print('Coletando dados...')
        


        t1e = frame.loc[0, 'TURNO 1 E']
        t1s = frame.loc[0, 'TURNO 1 S']
        t2e = frame.loc[0, 'TURNO 2 E']
        t2s = frame.loc[0, 'TURNO 2 S']
        t3e = frame.loc[0, 'TURNO 3 E']
        t3s = frame.loc[0, 'TURNO 3 S']
        
        ida = frame.loc[0, 'ID A'].astype(int)
        idb = frame.loc[0, 'ID B'].astype(int)
        idc = frame.loc[0, 'ID C'].astype(int)
        idd = frame.loc[0, 'ID D'].astype(int)
        ide = frame.loc[0, 'ID E'].astype(int)

        ano = frame.loc[0, 'ANO'].astype(int)
        mes = frame.loc[0, 'MÊS'].astype(int)
        print('Dados coletados!')
        print('Iniciando os filtros e as escritas nos arquivos de saída...')
        dia = 0
        for i in frame['ESCALA A'].dropna():
            dia=dia+1
            if (i == 1):
                arq_all.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(excecao)+"',"+str(t1e)+"','Exceção',"+str(t1s)+","+str(ida)+");"+'\n')
                
                arq = open (str(output)+'escala A mes-'+str(mes)+'.txt', 'a')
                arq.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(excecao)+"',"+str(t1e)+"','Exceção',"+str(t1s)+","+str(ida)+");"+'\n')
                arq.close ()
            
            if (i == 2):
                arq_all.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(excecao)+"',"+str(t2e)+",'Exceção',"+str(t2s)+","+str(ida)+");"+'\n')

                arq = open (str(output)+'escala A mes-'+str(mes)+'.txt', 'a')
                arq.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(excecao)+"',"+str(t2e)+",'Exceção',"+str(t2s)+","+str(ida)+");"+'\n')
                arq.close ()
            
            if (i == 3):
                arq_all.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(excecao)+"',"+str(t3e)+",'Exceção',"+str(t3s)+","+str(ida)+");"+'\n')

                arq = open (str(output)+'escala A mes-'+str(mes)+'.txt', 'a')
                arq.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(excecao)+"',"+str(t3e)+",'Exceção',"+str(t3s)+","+str(ida)+");"+'\n')
                arq.close ()
            
            if (i == 0):
                arq_all.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(excecao)+"','null','Exceção','null',"+str(ida)+");"+'\n')
                
                arq = open (str(output)+'escala A mes-'+str(mes)+'.txt', 'a')
                arq.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(excecao)+"','null','Exceção','null',"+str(ida)+");"+'\n')
                arq.close ()
        print('Arquivo criado com sucesso!!')

        dia=0
        for i in frame['ESCALA B'].dropna():    
            dia=dia+1
            if (i == 1):
                arq_all.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(excecao)+"',"+str(t1e)+",'Exceção',"+str(t1s)+","+str(idb)+");"+'\n')

                arq = open (str(output)+'escala B mes-'+str(mes)+'.txt', 'a')
                arq.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(excecao)+"',"+str(t1e)+",'Exceção',"+str(t1s)+","+str(idb)+");"+'\n')
                arq.close ()
            
            if (i == 2):
                arq_all.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(excecao)+"',"+str(t2e)+",'Exceção',"+str(t2s)+","+str(idb)+");"+'\n')

                arq = open (str(output)+'escala B mes-'+str(mes)+'.txt', 'a')
                arq.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(excecao)+"',"+str(t2e)+",'Exceção',"+str(t2s)+","+str(idb)+");"+'\n')
                arq.close ()
            
            if (i == 3):
                arq_all.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"',"+str(excecao)+"',"+str(t3e)+",'Exceção',"+str(t3s)+","+str(idb)+");"+'\n')

                arq = open (str(output)+'escala B mes-'+str(mes)+'.txt', 'a')
                arq.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(excecao)+"',"+str(t3e)+",'Exceção',"+str(t3s)+","+str(idb)+");"+'\n')
                arq.close ()
            
            if (i == 0):
                arq_all.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(excecao)+"','null','Exceção','null',"+str(ida)+");"+'\n')
                
                arq = open (str(output)+'escala A mes-'+str(mes)+'.txt', 'a')
                arq.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(excecao)+"','null','Exceção','null',"+str(ida)+");"+'\n')
                arq.close ()
        print('Arquivo criado com sucesso!!')

        dia=0
        for i in frame['ESCALA C'].dropna():    
            dia=dia+1
            if (i == 1):
                arq_all.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(excecao)+"',"+str(t1e)+",'Exceção',"+str(t1s)+","+str(idc)+");"+'\n')

                arq = open (str(output)+'escala C mes-'+str(mes)+'.txt', 'a')
                arq.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(excecao)+"',"+str(t1e)+",'Exceção',"+str(t1s)+","+str(idc)+");"+'\n')
                arq.close ()
            
            if (i == 2):
                arq_all.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(excecao)+"',"+str(t2e)+",'Exceção',"+str(t2s)+","+str(idc)+");"+'\n')

                arq = open (str(output)+'escala C mes-'+str(mes)+'.txt', 'a')
                arq.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(excecao)+"',"+str(t2e)+",'Exceção',"+str(t2s)+","+str(idc)+");"+'\n')
                arq.close ()
            
            if (i == 3):
                arq_all.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(excecao)+"',"+str(t3e)+",'Exceção',"+str(t3s)+","+str(idc)+");"+'\n')

                arq = open (str(output)+'escala C mes-'+str(mes)+'.txt', 'a')
                arq.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(excecao)+"',"+str(t3e)+",'Exceção',"+str(t3s)+","+str(idc)+");"+'\n')
                arq.close ()
            
            if (i == 0):
                arq_all.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(excecao)+"','null','Exceção','null',"+str(ida)+");"+'\n')
                
                arq = open (str(output)+'escala A mes-'+str(mes)+'.txt', 'a')
                arq.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(excecao)+"','null','Exceção','null',"+str(ida)+");"+'\n')
                arq.close ()
        print('Arquivo criado com sucesso!!')

        dia=0
        for i in frame['ESCALA D'].dropna():    
            dia=dia+1
            if (i == 1):
                arq_all.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(excecao)+"',"+str(t1e)+",'Exceção',"+str(t1s)+","+str(idd)+");"+'\n')

                arq = open (str(output)+'escala D mes-'+str(mes)+'.txt', 'a')
                arq.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(excecao)+"',"+str(t1e)+",'Exceção',"+str(t1s)+","+str(idd)+");"+'\n')
                arq.close ()
            
            if (i == 2):
                arq_all.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(excecao)+"',"+str(t2e)+",'Exceção',"+str(t2s)+","+str(idd)+");"+'\n')

                arq = open (str(output)+'escala D mes-'+str(mes)+'.txt', 'a')
                arq.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(excecao)+"',"+str(t2e)+",'Exceção',"+str(t2s)+","+str(idd)+");"+'\n')
                arq.close ()
            
            if (i == 3):
                arq_all.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(excecao)+"',"+str(t3e)+",'Exceção',"+str(t3s)+","+str(idd)+");"+'\n')

                arq = open (str(output)+'escala D mes-'+str(mes)+'.txt', 'a')
                arq.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(excecao)+"',"+str(t3e)+",'Exceção',"+str(t3s)+","+str(idd)+");"+'\n')
                arq.close ()
            
            if (i == 0):
                arq_all.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(excecao)+"','null','Exceção','null',"+str(ida)+");"+'\n')
                
                arq = open (str(output)+'escala A mes-'+str(mes)+'.txt', 'a')
                arq.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(excecao)+"','null','Exceção','null',"+str(ida)+");"+'\n')
                arq.close ()
        print('Arquivo criado com sucesso!!')

        dia=0
        for i in frame['ESCALA E'].dropna():
            dia=dia+1
            if (i == 1):
                arq_all.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(excecao)+"',"+str(t1e)+",'Exceção',"+str(t1s)+","+str(ide)+");"+'\n')

                arq = open (str(output)+'escala E mes-'+str(mes)+'.txt', 'a')
                arq.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(excecao)+"',"+str(t1e)+",'Exceção',"+str(t1s)+","+str(ide)+");"+'\n')
                arq.close ()
            
            if (i == 2):
                arq_all.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(excecao)+"',"+str(t2e)+",'Exceção',"+str(t2s)+","+str(ide)+");"+'\n')

                arq = open (str(output)+'escala E mes-'+str(mes)+'.txt', 'a')
                arq.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(excecao)+"',"+str(t2e)+",'Exceção',"+str(t2s)+","+str(ide)+");"+'\n')
                arq.close ()
            
            if (i == 3):
                arq_all.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(excecao)+"',"+str(t3e)+",'Exceção',"+str(t3s)+","+str(ide)+");"+'\n')

                arq = open (str(output)+'escala E mes-'+str(mes)+'.txt', 'a')
                arq.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(excecao)+"',"+str(t3e)+",'Exceção',"+str(t3s)+","+str(ide)+");"+'\n')
                arq.close ()
            
            if (i == 0):
                arq_all.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(excecao)+"','null','Exceção','null',"+str(ida)+");"+'\n')
                
                arq = open (str(output)+'escala A mes-'+str(mes)+'.txt', 'a')
                arq.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"','"+str(excecao)+"','null','Exceção','null',"+str(ida)+");"+'\n')
                arq.close ()

            
        arq_all.close()
        print('Arquivo criado com sucesso!!')
        print('Concluído!')

        return True