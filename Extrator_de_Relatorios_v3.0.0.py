from csv import Dialect
from os import read, terminal_size
from ttkbootstrap import Style
import pymysql.cursors
import pandas as pd
from tkinter import *
from tkcalendar import *
from datetime import date, timedelta
from tkinter import messagebox
import locale
import threading
from tkinter import ttk, messagebox
from datetime import timedelta, date
import xlsxwriter
from tkcalendar import DateEntry  # Verificar a versão do pacote 'tkcalendar'

# Define o idioma
locale.setlocale(locale.LC_ALL, 'pt_BR.utf8')

# conexão BD
con = pymysql.connect(
    host='mysql01.socicam.com.br',
    port=3307,
    user='report',
    password='Yo5WOMSOnKHh',
    database='embarqueweb',
    cursorclass=pymysql.cursors.DictCursor
)

con2 = pymysql.connect(
    host='mysql01.socicam.com.br',
    port=3307,
    user='report',
    password='Yo5WOMSOnKHh',
    database='socicam',
    cursorclass=pymysql.cursors.DictCursor
)


def gerar_relatorio():
    # Cursor metodo .cursor()
    with con.cursor()as c:

        # Pega as datas de inicio e fim do calendario e adiciona no select como uma Fstring
        dataInicial = calendario.get_date()
        dataFinal = calendario.get_date()

        validacao = dataFinal - dataInicial
        if validacao >= timedelta(days=186):
            messagebox.showinfo(
                'Erro', 'O período máximo para consulta é de 6 meses')
            raise Exception

        global unidades_norte
        global unidades_sul
        global unidades_fortaleza
        global unidades_controladoria

        unidades_norte = ''
        unidades_sul = ''
        unidades_fortaleza = ''
        unidades_controladoria = ''
        if var_divNorte.get() == 1:
            unidades_norte = 2,4,5,6,8,9,10,12,13,14,15,16,17,19,20,21,22,23,24,25,34
        elif var_divSul.get() == 1:
            unidades_sul = 33, 7, 11, 3, 31, 1, 27, 41
        elif var_fortaleza.get() == 1:
            unidades_fortaleza = 8, 9, 10


        # Consulta no BD
        sql_1 = f"""
SELECT
t.txtfantasyname AS Terminal,
e.txtcorporatename AS RazaoSocialEmp,
e.txtfantasyname AS NomeFantasiaEmp,
e.txtcnpj AS CNPJEmp,
s.numcdt AS NumeroTermo,
r.numreceipt AS NumeroRecibo,
r.datereceipt AS DataRecibo,
p.txttype AS FormaRecebimento,
Coalesce(ti.txtabbreviation, '*') AS TipoTarifa,
COALESCE(SUM(IF(sr.numTicketReceived >= cdt_tickets.qttTickets, cdt_tickets.qttTickets - cdt_tickets.cancelled, IF(s.ticket_id IS NOT NULL, sr.numTicketReceived, IF(cdt_tickets.qttTickets - cdt_tickets.cancelled > 0, cdt_tickets.qttTickets - cdt_tickets.cancelled, 0)))), 0) AS QtdeRecebida,
SUM(IF(s.flagIsCDT AND cdt_tickets.qttTickets - cdt_tickets.cancelled > 0, (sr.valueReceived / sr.numTicketReceived) * (cdt_tickets.qttTickets - cdt_tickets.cancelled), sr.valueReceived)) AS ValorTotalRecebido
FROM embarqueweb.receipts r
INNER JOIN sale_receipt sr ON sr.receipt_id = r.id
INNER JOIN sales s ON s.id = sr.sale_id
INNER JOIN terminals t ON t.id = r.terminal_id
LEFT JOIN embarqueweb.cdt_tickets ON cdt_tickets.sale_id = s.id AND cdt_tickets.qttTickets - cdt_tickets.cancelled > 0 AND s.ticket_id IS NULL AND s.flagIsCDT
INNER JOIN embarqueweb.tickets ti ON ti.id = COALESCE(s.ticket_id, cdt_tickets.ticket_id)
INNER JOIN enterprises e ON e.id = r.enterprise_id
INNER JOIN paymenttypes p ON p.id = r.paymenttype_id
WHERE r.terminal_id IN {unidades_norte}{unidades_sul}{unidades_fortaleza}
AND r.flagstatus
AND s.flagstatus
AND r.datereceipt BETWEEN '{dataInicial} 00:00:00' AND '{dataFinal} 23:59:59'
GROUP BY
t.txtfantasyname,
e.txtfantasyname,
e.txtcnpj,
s.numcdt,
r.numreceipt,
r.datereceipt,
p.txttype,
ti.txtabbreviation
ORDER BY
t.txtfantasyname,
r.datereceipt,
ti.txtabbreviation;
        """
        #FORMAT(rps.vlrRecebido, 2)
        c.execute(sql_1)  # CONVERT(money,rps.vlrRecebido)
        resultado1 = c.fetchall()
        print(resultado1) #2,4,5,6,8,9,10,12,13,14,15,16,17,19,20,21,22,23,24,25,34

    # Percorre a lista e o dict e transforma VlrRecebido em float
    for dicts in resultado1:
        for keys in dicts:
            dicts['ValorTotalRecebido'] = float(dicts['ValorTotalRecebido'])
            # float(dicts['ValorTotalRecebido'])

    # Percorre a lista e o dict e transforma QtdeRecebida em int
    for dicts in resultado1:
        for keys in dicts:
            dicts['QtdeRecebida'] = int(dicts['QtdeRecebida'])

    print(resultado1)
    relatorio = resultado1
    relatorio = pd.DataFrame(relatorio)
    writer = pd.ExcelWriter(f'Relatorio_Recebimento_{dataInicial}__{dataFinal}.xlsx', engine='xlsxwriter',
                            engine_kwargs={'options': {'strings_to_float': True}})
    relatorio.to_excel(writer, sheet_name='recebimento',
                       index=False, float_format="%.2f")
    workbook = writer.book
    worksheet = writer.sheets['recebimento']
    worksheet.autofilter('A1:K1')  # Adiciona o Filtro

    # Tamanho das Celulas
    worksheet.set_column('A:A', 39)
    worksheet.set_column('B:B', 49)
    worksheet.set_column('C:C', 40)
    worksheet.set_column('D:K', 20)
    # R$ #,###.00
    format_number = workbook.add_format({'num_format': '#R$ #,###.00'})
    worksheet.set_column('K1:K10000', 20, format_number)

    writer.save()
    c.close()
    #salvar = filedialog.asksaveasfile()


def msg_conclusao():
    return messagebox.showinfo('Relatório Recebimento', 'Relatório gerado')

def msg_entrega():
    return messagebox.showinfo('Relatório Entrega', 'Relatório de Entrega gerado')

'''def selectDate():
    dataInicial = calendario.get_date()
    selectDate = Label(text=dataInicial)
    selectDate.place(x = 250, y = 250)
'''

# Interface Gráfica
janela = Tk()
style = Style()
janela = style.master
janela.title('Extrator de relatórios Excel')
janela.configure(background='#A1CDEC')
janela.iconbitmap('imagens\Socicam.ico')
janela.resizable(False, False)
#janela.minsize(280, 300)
#Tamanho da Janela
window_width = 420
window_height = 450

# Centraliza a tela
screen_width = janela.winfo_screenwidth()
screen_height = janela.winfo_screenheight()
center_x = int(screen_width/2 - window_width / 2)
center_y = int(screen_height/2 - window_height / 2)
janela.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')


# Adiciona Imagem
img = PhotoImage(file='imagens\logo-socicam.png')
label_imagem = Label(janela, image=img)
label_imagem.configure(background='#A1CDEC')
label_imagem.place(relx=0.5, rely=0.8, anchor=CENTER)
                                #0.7

texto = Label(janela, text='Período do Relatório:',
              font=('calibri', 15, 'bold'))
texto.configure(background='#A1CDEC')
texto.pack(ipady=7)


label_versao = Label(janela, text='v3.0.0 - Socicam',
                     font=('calibri', 8, 'italic'))
label_versao.place(x=1, y=430)
label_versao.configure(background='#A1CDEC')


def all_commands(): 
    return [gerar_relatorio(), msg_conclusao()]

def all_entrega():
    return [entrega(), msg_entrega()]


def gerar_embarques():
    with con.cursor() as c2:
        #Datas do periodo
        dataInicial = calendario.get_date()
        dataFinal = calendario.get_date()
        #Validação da Consulta-Catracas
        validacao = dataFinal - dataInicial
        if validacao >= timedelta(days = 31):
            messagebox.showinfo('Erro', 'O período máximo para consulta é de 1 mês')
            raise Exception
        global unidades_norte
        global unidades_sul
        global unidades_fortaleza
        global unidades_controladoria

        unidades_norte = ''
        unidades_sul = ''
        unidades_fortaleza = ''
        unidades_controladoria = ''
        if var_divNorte.get() == 1:
            unidades_norte = 2,4,5,6,8,9,10,12,13,14,15,16,17,19,20,21,22,23,24,25,34
        elif var_divSul.get() == 1:
            unidades_sul = 33, 7, 11, 3, 31, 1, 27, 41
        elif var_fortaleza.get() == 1:
            unidades_fortaleza = 8, 9, 10
        #Consulta no BD
        sql_2 = f"""SELECT terminals.txtFantasyName AS Terminal,
                    DATE(numerations.dateReadTurnstile) AS `Data`,
                    numerations.codeTurnstile as Catraca,
                    COUNT(numerations.id) AS Embarques
                    FROM embarqueweb.numerations
                    INNER JOIN embarqueweb.sales ON sales.id = numerations.sale_id
                    INNER JOIN embarqueweb.enterprises ON sales.enterprise_id = enterprises.id
                    INNER JOIN embarqueweb.terminals ON sales.terminal_id = terminals.id
                    LEFT JOIN embarqueweb.numerations_ticket ON numerations_ticket.numeration_id = numerations.id
                    LEFT JOIN tickets ti ON ti.id = sales.ticket_id
                    WHERE 1=1
                    AND sales.flagStatus
                    AND terminals.id in {unidades_norte}{unidades_sul}{unidades_fortaleza}
                    AND numerations.dateReadTurnstile between '{dataInicial} 00:00:00' AND '{dataFinal} 23:59:59'
                    GROUP BY terminals.txtFantasyName, date(numerations.dateReadTurnstile), numerations.codeTurnstile
                    ORDER BY terminals.txtFantasyName, date(numerations.dateReadTurnstile), numerations.codeTurnstile;
                """
        c2.execute(sql_2)
        resultado2 = c2.fetchall()
        print(resultado2)
        #Gerando o Excel
        relatorio_catracas = resultado2
        relatorio_catracas = pd.DataFrame(relatorio_catracas)
        writer = pd.ExcelWriter(f'Relatorio_Embarques_{dataInicial}__{dataFinal}.xlsx', engine = 'xlsxwriter')
        relatorio_catracas.to_excel(writer, sheet_name = 'Catracas', index = False)
        workbook = writer.book
        worsheet = writer.sheets['Catracas']
        worsheet.autofilter('A1:D1') #Adiciona o Filtro

        #Tamanho das Celulas
        worsheet.set_column('A:A', 39)
        worsheet.set_column('B:B', 49)
        worsheet.set_column('C:C', 20)
        worsheet.set_column('D:D', 15)
        writer.save()
        c2.close()
        return messagebox.showinfo('Relatório Embarques','Relatório gerado'),barra1.stop(),popup1.quit()


def gerar_estatistico():
    # Consulta Vendas
    with con.cursor() as c3:
        # Datas do periodo
        dataInicial = calendario.get_date()
        dataFinal = calendario.get_date()
        validacao = dataFinal - dataInicial
        if validacao >= timedelta(days = 31):
            messagebox.showinfo('Erro', 'O período máximo para consulta é de 1 mês')
            raise Exception
        #Consulta no BD
        sql_vendas = f"""SELECT t.txtFantasyName AS Terminal,YEAR(s.dateSale) AS Ano,MONTH(s.dateSale) AS Mes,DAY(s.dateSale) AS Dia,ti.txtAbbreviation AS TipoTarifa,
                            (CASE WHEN (cdt.qttTickets IS NOT NULL AND s.terminal_id <> 17)
                            THEN SUM(cdt.qttTickets) ELSE SUM(s.numQuantityRequested) END) AS QtdeVendida
                            FROM sales s
                            LEFT JOIN cdt_tickets cdt ON cdt.sale_id = s.id,
                            tickets ti,
                            terminals t
                            WHERE s.terminal_id = t.id
                            AND ((s.ticket_id IS NOT NULL AND s.ticket_id = ti.id) OR (s.ticket_id IS NULL AND cdt.ticket_id = ti.id))
                            AND s.flagStatus IS true
                            AND s.dateSale BETWEEN '{dataInicial} 00:00:00' AND '{dataFinal} 23:59:59'
                            AND s.enterprise_id <> 305
                            AND s.terminal_id <> 35
                            GROUP BY
                            t.txtFantasyName,
                            YEAR(s.dateSale),
                            MONTH(s.dateSale),
                            DAY(s.dateSale),
                            ti.txtAbbreviation
                            ORDER BY
                            t.txtFantasyName,
                            YEAR(s.dateSale),
                            MONTH(s.dateSale),
                            DAY(s.dateSale),
                            ti.txtAbbreviation
                            LIMIT 1000000;
                            """
        c3.execute(sql_vendas)
        resultado_vendas = c3.fetchall()
        for dicts in resultado_vendas:
            for keys in dicts:
                dicts['QtdeVendida'] = float(dicts['QtdeVendida'])
        print(resultado_vendas)
    # Consulta Recebimentos
    with con.cursor() as c4:
        # Datas do Periodo
        dataInicial = calendario.get_date()
        datFinal = calendario.get_date()
        validacao = dataFinal - dataInicial
        if validacao >= timedelta(days = 31):
            messagebox.showinfo('Erro', 'O período máximo para consulta é de 1 mês')
            raise Exception
        # Consulta no BD
        sql_recebimentos = f""" SELECT  t.txtFantasyName AS Terminal,
                                YEAR(r.dateReceipt) AS Ano,
                                MONTH(r.dateReceipt) AS Mes,
                                DAY(r.dateReceipt) AS Dia,
                                COALESCE(ti.txtAbbreviation, '*') AS TipoTarifa,
                                sum(sr.numTicketReceived) as QtdeRecebida
                                FROM embarqueweb.receipts r
                                INNER JOIN sale_receipt sr ON sr.receipt_id = r.id
                                INNER JOIN sales s ON s.id = sr.sale_id
                                INNER JOIN terminals t ON t.id = r.terminal_id
                                LEFT JOIN tickets ti ON ti.id = s.ticket_id
                                INNER JOIN enterprises e ON e.id = r.enterprise_id
                                INNER JOIN paymenttypes p ON p.id = r.paymenttype_id
                                WHERE 1=1 -- r.terminal_id in (31)
                                AND r.flagStatus = true
                                AND s.flagStatus = true
                                AND r.dateReceipt BETWEEN '{dataInicial} 00:00:00' AND '{dataFinal} 23:59:59'
                                GROUP BY
                                t.txtFantasyName,
                                YEAR(r.dateReceipt),
                                MONTH(r.dateReceipt),
                                DAY(r.dateReceipt),
                                ti.txtAbbreviation
                                ORDER BY
                                t.txtFantasyName,
                                YEAR(r.dateReceipt),
                                MONTH(r.dateReceipt),
                                DAY(r.dateReceipt),
                                ti.txtAbbreviation
                                LIMIT 2000000;
                            """
        c4.execute(sql_recebimentos)
        resultado_recebimentos = c4.fetchall()
        for dicts in resultado_recebimentos:
            for keys in dicts:
                dicts['QtdeRecebida'] = int(dicts['QtdeRecebida'])
        print(resultado_recebimentos)
    # Consulta Embarques - Emissor Nautico
    with con2.cursor() as c5:
        # Datas do Periodo
        dataInicial = calendario.get_date()
        datFinal = calendario.get_date()
        validacao = dataFinal - dataInicial
        if validacao >= timedelta(days = 31):
            messagebox.showinfo('Erro', 'O período máximo para consulta é de 1 mês')
            raise Exception
        #Consulta no BD 'Sistema Emissor Náutico' as sistema
        sql_embarque_nautico = f"""SELECT 'TERMINAL NAUTICO DA BAHIA' as Terminal,
                                YEAR(DATE_ADD(s.turnstile_datetime, INTERVAL s.turnstile_datetime_timezone HOUR)) AS Ano,
                                MONTH(DATE_ADD(s.turnstile_datetime, INTERVAL s.turnstile_datetime_timezone HOUR)) AS Mes,
                                DAY(DATE_ADD(s.turnstile_datetime, INTERVAL s.turnstile_datetime_timezone HOUR)) AS Dia,
                                'UN' as TipoTarifa,
                                COUNT(s.id) as QtdeEmbarcada
                                FROM socicam.shipping s
                                WHERE
                                DATE_ADD(s.turnstile_datetime, INTERVAL s.turnstile_datetime_timezone HOUR) 
                                between '{dataInicial} 00:00:00' AND '{dataFinal} 23:59:59'
                                GROUP BY
                                YEAR(DATE_ADD(s.turnstile_datetime, INTERVAL s.turnstile_datetime_timezone HOUR)),
                                MONTH(DATE_ADD(s.turnstile_datetime, INTERVAL s.turnstile_datetime_timezone HOUR)),
                                DAY(DATE_ADD(s.turnstile_datetime, INTERVAL s.turnstile_datetime_timezone HOUR))
                                ORDER BY
                                YEAR(DATE_ADD(s.turnstile_datetime, INTERVAL s.turnstile_datetime_timezone HOUR)),
                                MONTH(DATE_ADD(s.turnstile_datetime, INTERVAL s.turnstile_datetime_timezone HOUR)),
                                DAY(DATE_ADD(s.turnstile_datetime, INTERVAL s.turnstile_datetime_timezone HOUR))
                                LIMIT 1000000;
                            """
        c5.execute(sql_embarque_nautico)
        resultado_embarque_nautico = c5.fetchall()
        print(resultado_embarque_nautico)
    with con.cursor() as c6:
        # Datas do Periodo
        dataInicial = calendario.get_date()
        datFinal = calendario.get_date()
        validacao = dataFinal - dataInicial
        if validacao >= timedelta(days = 31):
            messagebox.showinfo('Erro', 'O período máximo para consulta é de 1 mês')
            raise Exception
        # Consulta no BD
        sql_embarques = f"""SELECT terminals.txtFantasyName AS Terminal,
                                year(numerations.dateReadTurnstile) AS `Ano`,
                                month(numerations.dateReadTurnstile) AS `Mês`,
                                DAy(numerations.dateReadTurnstile) AS `Dia`,
                                COALESCE(ti.txtAbbreviation, '*') AS ticketType,
	                            COUNT(numerations.id) AS Embarques
                                FROM embarqueweb.numerations
                                INNER JOIN embarqueweb.sales ON sales.id = numerations.sale_id
                                INNER JOIN embarqueweb.terminals ON sales.terminal_id = terminals.id
                                LEFT JOIN embarqueweb.numerations_ticket ON numerations_ticket.numeration_id = numerations.id
                                LEFT JOIN tickets ti ON ti.id = sales.ticket_id
                                WHERE 1=1
                                AND sales.flagStatus
                                AND numerations.dateReadTurnstile BETWEEN '{dataInicial} 00:00:00' AND '{dataFinal} 23:59:59'
                                GROUP BY terminals.txtFantasyName, year(numerations.dateReadTurnstile),month(numerations.dateReadTurnstile),DAy(numerations.dateReadTurnstile), ticketType
                                ORDER BY terminals.txtFantasyName, year(numerations.dateReadTurnstile),month(numerations.dateReadTurnstile),Day(numerations.dateReadTurnstile);
                         """
        c6.execute(sql_embarques)
        resultado_embarques = c6.fetchall()
        print(resultado_embarques)
    #Variaveis relatorio para gerar o Excel
    relatorio_vendas = resultado_vendas                #{'Data': [11, 12, 13, 14]}
    relatorio_recebimentos = resultado_recebimentos    #{'Data': [21, 22, 23, 24]}
    relatorio_nautico = resultado_embarque_nautico     #{'Data': [31, 32, 33, 34]}
    relatorio_embarques = resultado_embarques
    #emissor_nautico = {'Emissor Náutico':['Sistema Emissor Náutico']}
    #adiciona_nautico = emissor_nautico
    #Gerando Relatorios Excel
    relatorio_vendas = pd.DataFrame(relatorio_vendas)
    relatorio_recebimentos = pd.DataFrame(relatorio_recebimentos)
    relatorio_nautico = pd.DataFrame(relatorio_nautico)
    #adiciona_nautico = pd.DataFrame(adiciona_nautico)
    relatorio_embarques = pd.DataFrame(relatorio_embarques)
    writer = pd.ExcelWriter(f'Estatistico-Sistema-Embarque_{dataInicial}__{dataFinal}.xlsx', engine='xlsxwriter')
    relatorio_vendas.to_excel(writer, sheet_name='Vendas', index= False)
    relatorio_recebimentos.to_excel(writer, sheet_name='Recebimentos', index= False)
    relatorio_embarques.to_excel(writer,sheet_name='Embarques(Catraca)', index=False)
    relatorio_nautico.to_excel(writer, sheet_name='Emissor Náutico', index= False)
    #adiciona_nautico.to_excel(writer, sheet_name='Emissor Náutico', startcol=6, index=False)
    # Nome das Sheets
    workbook = writer.book
    worksheet_vendas = writer.sheets['Vendas']
    worksheet_recebimentos = writer.sheets['Recebimentos']
    worksheet_estatistico = writer.sheets['Embarques(Catraca)']
    worksheet_nautico = writer.sheets['Emissor Náutico']
    # Adiciona os Filtros
    worksheet_vendas.autofilter('A1:F1')
    worksheet_recebimentos.autofilter('A1:F1')
    worksheet_estatistico.autofilter('A1:F1')
    worksheet_nautico.autofilter('A1:F1')
    # Tamanho das Colunas
    # Tamanho Vendas
    worksheet_vendas.set_column('A:A', 39)
    worksheet_vendas.set_column('B:B', 8)
    worksheet_vendas.set_column('C:C', 9)
    worksheet_vendas.set_column('D:D', 8)
    worksheet_vendas.set_column('E:E', 14)
    worksheet_vendas.set_column('F:F', 18)
    # Tamanho Recebimentos
    worksheet_recebimentos.set_column('A:A', 39)
    worksheet_recebimentos.set_column('B:B', 8)
    worksheet_recebimentos.set_column('C:C', 9)
    worksheet_recebimentos.set_column('D:D', 8)
    worksheet_recebimentos.set_column('E:E', 14)
    worksheet_recebimentos.set_column('F:F', 18)
    # Header Coluna
    header = 'TipoTarifa'
    bold = workbook.add_format({'bold': True})
    worksheet_estatistico.write('E1', header, bold)


    # Tamanho Estatistico
    worksheet_estatistico.set_column('A:A', 39)
    worksheet_estatistico.set_column('B:B', 8)
    worksheet_estatistico.set_column('C:C', 9)
    worksheet_estatistico.set_column('D:D', 8)
    worksheet_estatistico.set_column('E:E', 14)
    worksheet_estatistico.set_column('F:F', 18)
    worksheet_estatistico.set_column('G:G', 23)
    # Tamanho Nautico
    worksheet_nautico.set_column('A:A', 39)
    worksheet_nautico.set_column('B:B', 8)
    worksheet_nautico.set_column('C:C', 9)
    worksheet_nautico.set_column('D:D', 8)
    worksheet_nautico.set_column('E:E', 14)
    worksheet_nautico.set_column('F:F', 18)
    #worksheet_nautico.set_column('G:G', 23)


    writer.save()
    c3.close()
    c4.close()
    c5.close()
    c6.close()
    return messagebox.showinfo('Relatório Estatístico','Relatório gerado'),barra1.stop(),popup1.quit()




def entrega():
    with con.cursor() as c7:
        dataInicial = calendario.get_date()
        dataFinal = calendario.get_date()
        validacao = dataFinal - dataInicial
        if validacao >= timedelta(days=186):
            messagebox.showinfo(
                'Erro', 'O período máximo para consulta é de 6 meses')
            raise Exception
        global unidades_norte
        global unidades_sul
        global unidades_fortaleza
        global unidades_controladoria

        unidades_norte = ''
        unidades_sul = ''
        unidades_fortaleza = ''
        unidades_controladoria = ''
        unidades_norte_sul = ''
        if var_divNorte.get() == 1:
            unidades_norte = 2,4,5,6,8,9,10,12,13,14,15,16,17,19,20,21,22,23,24,25,34
        elif var_divSul.get() == 1:
            unidades_sul = 33, 7, 11, 3, 31, 1, 27, 41
        elif var_fortaleza.get() == 1:
            unidades_fortaleza = 8, 9, 10

        sql_entrega = f""" 
        SELECT 	t.txtFantasyName AS Terminal, 
		s.numCDT AS numCDT, 
        s.numCdtBlock AS numCDTBloco, 
        e.txtFantasyName AS Empresa, 
        e.txtCNPJ AS CNPJEmpresa, 
		ti.txtAbbreviation AS TipoDaTarifa, 
		(CASE WHEN (cdt.qttTickets IS NOT NULL) THEN cdt.valueUnit ELSE s.valueSaleUnit END) AS ValorUnitarioTarifa, 
		(CASE WHEN (cdt.qttTickets IS NOT NULL) THEN cdt.qttTickets ELSE s.numQuantityRequested END) AS QtdeTarifaSolicitada,
        (CASE WHEN (s.discount IS TRUE) THEN s.numTicketsDiscount ELSE '0.00' END) AS `Qtde Desconto`,
		(CASE WHEN (s.discount IS TRUE) THEN (CASE WHEN (cdt.qttTickets IS NOT NULL) THEN (cdt.valueUnit * (cdt.qttTickets - s.numTicketsDiscount)) ELSE s.valueTotalSale - (s.numTicketsDiscount * s.valueSaleUnit) END)
        ELSE
        (CASE WHEN (cdt.qttTickets IS NOT NULL) THEN (cdt.valueUnit * cdt.qttTickets) ELSE s.valueTotalSale END) END) AS TotalVenda,
		s.dateSale AS DataDaVenda, 
        s.dateDue AS DataDoVencimento
FROM 	sales s LEFT JOIN cdt_tickets cdt ON cdt.sale_id = s.id, 
		tickets ti, 
        terminals t,
        enterprises e
WHERE 	s.terminal_id = t.id 
		AND s.terminal_id IN {unidades_norte}{unidades_sul}{unidades_fortaleza}{unidades_norte_sul} /* ID do terminal (33, 7, 11, 3, 31, 1, 27) */
        AND s.enterprise_id <> 305
		AND s.enterprise_id = e.id
		AND ((s.ticket_id IS NOT NULL AND s.ticket_id = ti.id) OR (s.ticket_id IS NULL AND cdt.ticket_id = ti.id))
		AND s.flagStatus IS TRUE
		AND s.dateSale BETWEEN '{dataInicial} 00:00:00' AND '{dataFinal} 23:59:59' /* Intervalo de datas */
ORDER BY 
		t.txtFantasyName, 
        s.dateSale
LIMIT 99999999;
         """
        c7.execute(sql_entrega)
        resultado_entrega = c7.fetchall()
        print(resultado_entrega)

        for dicts in resultado_entrega:
            for keys in dicts:
                dicts['ValorUnitarioTarifa'] = float(dicts['ValorUnitarioTarifa'])
        
        for dicts in resultado_entrega:
            for keys in dicts:
                dicts['Qtde Desconto'] = float(dicts['Qtde Desconto'])

        for dicts in resultado_entrega:
            for keys in dicts:
                dicts['TotalVenda'] = float(dicts['TotalVenda'])

        relatorio_entrega = resultado_entrega
        relatorio_entrega = pd.DataFrame(relatorio_entrega)
        writer = pd.ExcelWriter(f'Entregas_Detalhado_{dataInicial}__{dataFinal}.xlsx', engine='xlsxwriter')
        relatorio_entrega.to_excel(writer, sheet_name='Entregas', index=False, float_format='%.2f')
        workbook = writer.book
        worksheet = writer.sheets['Entregas']
        # Adiciona Filtro
        worksheet.autofilter('A1:K1')
        # Tamanho das Celulas
        worksheet.set_column('A:A', 39)
        worksheet.set_column('B:B', 13)
        worksheet.set_column('C:C', 16)
        worksheet.set_column('D:D', 39)
        worksheet.set_column('E:E', 17)
        worksheet.set_column('F:F', 15)
        worksheet.set_column('G:G', 21)
        worksheet.set_column('H:H', 17)
        worksheet.set_column('I:I', 14)
        worksheet.set_column('J:J', 19)
        worksheet.set_column('K:K', 18)
        worksheet.set_column('L:L', 18)
        #Adiciona formato Reais
        format_number = workbook.add_format({'num_format': '#R$ #,###.00'})
        worksheet.set_column('K1:K10000', 21, format_number)
        writer.save()
        c7.close()





#def loading():
    #top = Toplevel(background='#A1CDEC')
    #top.title("Gerando Relatório")
    #top.resizable(False, False)
    #img = PhotoImage(file='imagens\loading.gif')
    #label_popup = Label(top, image=img)
    #label_popup.place(relx=0.5, rely=0.1, anchor=CENTER)

def disable_event():
    pass


def barra_carregamento():
    global barra1
    global popup1
    popup1 = Toplevel()
    label = Label(popup1, text='Gerando Relatório').pack()
    #popup1.iconbitmap('imagens\Socicam.ico')
    #popup1.resizable(False, False)
    #popup1.protocol('WM_DELETE_WINDOW', disable_event)
    popup1.overrideredirect(1)#Remove todos os botões
    # Centraliza a tela
    window_width = 110
    window_height = 50
    screen_width = janela.winfo_screenwidth()
    screen_height = janela.winfo_screenheight()
    center_x = int(screen_width/2 - window_width / 2)
    center_y = int(screen_height/2 - window_height / 2)
    popup1.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
    barra1 = ttk.Progressbar(popup1,orient= HORIZONTAL,mode='indeterminate', length=120)
    barra1.pack()
    barra1.start(10)


def start_embarques_thread(event):
    global embarques_thread
    global popup
    #global barra
    embarques_thread = threading.Thread(target= gerar_embarques)
    embarques_thread.daemon = True
    popup = Toplevel()
    label = Label(popup, text='Gerando Relatório').pack()
    barra = ttk.Progressbar(popup,orient= HORIZONTAL,mode='indeterminate', length=120)
    barra.pack()
    barra.start(10)
    embarques_thread.start()
    popup.after(20, check_embarques_thread)

def check_embarques_thread():
    if embarques_thread.is_alive():
        popup.after(20, check_embarques_thread)
    else:
        barra1.stop()




def msg_conclusao_embarque():
    return messagebox.showinfo('Relatório Embarques','Relatório gerado'),barra1.stop(),popup1.quit()



def command_embarques(): 
    carregamento = threading.Thread(target= barra_carregamento)
    embarques = threading.Thread(target= gerar_embarques)
    return [carregamento.start(), embarques.start()]#barra_carregamento(),gerar_embarques(),msg_conclusao_embarque()
    
    
    
    
    #if embarques.is_alive():
        #popup1.after(20,check_embarques_thread)
    #else:
        #barra1.stop()



def estatistico_all():
    carregamento = threading.Thread(target= barra_carregamento)
    estatistico = threading.Thread(target= gerar_estatistico)
    return [carregamento.start(), estatistico.start()]

#Escolha dos radioButton-
def selecao_radio():
    va = var_radio.get()
    if va == 1:
        all_commands()
    elif va == 2:
        command_embarques()
    elif va == 3:
        estatistico_all()
    elif va == 5:
        all_entrega()



botao = ttk.Button(janela, text='Gerar Relatório',
               style='selectbg.Outline.TButton', 
               command=selecao_radio)
botao.place(relx=0.5, rely=0.95, anchor=CENTER)



var_radio = IntVar()

mycolor = '#A1CDEC'
radio_style = ttk.Style()
radio_style.configure('Wild.TRadiobutton', background = mycolor)

# Label tipos de relatorios
label_tipos = Label(janela, text='Tipos de Relatórios:')
label_tipos.configure(background='#A1CDEC', font=('calibri', 15, 'bold'))
label_tipos.place(relx= 0.5, rely=0.32, anchor=CENTER)


radio_entrega = ttk.Radiobutton(
    janela,
    text='Entrega',
    style='Wild.TRadiobutton',
    value= 5,
    variable= var_radio)
radio_entrega.place(relx=0.17, rely=0.4, anchor='w')


radio_recebimento = ttk.Radiobutton(
    janela,
    text='Recebimento',
    style= 'Wild.TRadiobutton',
    value = 1,
    variable = var_radio)
radio_recebimento.place(relx= 0.5, rely=0.4, anchor=CENTER)      #, anchor= 'w' padx=150, pady=10 .pack(pady=40) #place(relx=0.5, rely=0.3, anchor=CENTER)
# 0.7 - 0.8 - 0.5  0.3


radio_embarques = ttk.Radiobutton(
    janela, 
    text='Embarques',
    style= 'Wild.TRadiobutton', 
    value = 2,
    variable = var_radio)
radio_embarques.place(relx= 0.65, rely=0.4, anchor='w') # relx = 0.35  anchor='w' .place(relx=0.5, rely=0.4, anchor=CENTER)


#radio_fortaleza = ttk.Radiobutton(
    #janela,
    #text ='Fortaleza(Recebimento)',
    #style ='Wild.TRadiobutton',
    #value = 4,
    #variable= var_radio)
#radio_fortaleza.place(relx= 0.35, rely=0.5, anchor='w') # relx = 0.57  .place(relx=0.5, rely=0.6, anchor=CENTER)

radio_estatistico = ttk.Radiobutton(
    janela,
    text ='Controladoria',
    style = 'Wild.TRadiobutton',
    value = 3,
    variable= var_radio)
radio_estatistico.place(relx= 0.17, rely= 0.5, anchor='w')#.pack(side = 'left',fill=X)# side = 'top',anchor= 'w'.place(relx=0.5, rely= 0.5, anchor=CENTER)

# Estilo do CheckButton
checkbox_style = ttk.Style()
checkbox_style.configure('success.TCheckbutton', background = mycolor) #,foreground=[('active', 'red')])

label_unidades = Label(janela, text='Unidades:')
label_unidades.configure(background='#A1CDEC', font=('calibri', 15, 'bold'))
label_unidades.place(relx=0.5, rely=0.6, anchor=CENTER)

#Checkbuttons das Unidades

var_divNorte = IntVar()
var_divSul = IntVar()
var_fortaleza = IntVar()


checkbox_divNorte = ttk.Checkbutton(janela, text= 'Divisão Norte', style='success.TCheckbutton',variable = var_divNorte, onvalue = 1, offvalue = 0)
checkbox_divNorte.place(relx=0.1, rely= 0.7, anchor= 'w')

checkbox_divSul = ttk.Checkbutton(janela, text='Divisão Sul', style='success.TCheckbutton',variable = var_divSul, onvalue = 1, offvalue = 0)
checkbox_divSul.place(relx= 0.5, rely= 0.7, anchor= CENTER)

checkbox_fortaleza = ttk.Checkbutton(janela, text='Fortaleza', style='success.TCheckbutton',variable = var_fortaleza, onvalue = 1, offvalue = 0)
checkbox_fortaleza.place(relx= 0.67,rely= 0.7, anchor='w')


# Desabilita o radiobutton
#radio_estatistico.configure(state= DISABLED)


# Criação da janela principal
def calendario():
    janela = Tk()
    janela.title("Minha Janela")
    janela.geometry("800x600")

    texto_periodo = Label(janela, text='a')
    texto_periodo.configure(background='#A1CDEC', font=('calibri', 10, 'bold'))
    texto_periodo.place(relx=0.48, rely=0.16)

    data_minima = date(2021, 7, 1)
    calendario = DateEntry(janela, day=1, year=2022, setmode='day',
                           date_pattern='dd/mm/yyyy',
                           width=12,
                           background='darkblue',
                           locale='pt_BR.utf8')
    calendario.place(relx=0.19, rely=0.15)

    calendario2 = DateEntry(janela, year=2022,
                            setmode='day',
                            date_pattern='dd/mm/yyyy',
                            background='darkblue',
                            locale='pt_BR.utf8')
    calendario2.place(relx=0.53, rely=0.15)
   


# Validação de Úsuario
try:
    validacao_usuario = open('settings.txt','r')
except FileNotFoundError:
    messagebox.showwarning('Erro','Arquivo de configuração inexistente')

abrir = validacao_usuario.read()


# Validação Nordeste
if abrir == '10':
    #radio_fortaleza.configure(state=DISABLED)
    radio_estatistico.configure(state=DISABLED)
    checkbox_divSul.configure(state=DISABLED)
    checkbox_fortaleza.configure(state=DISABLED)
    # Validação Fortaleza
elif abrir == '20':
    radio_estatistico.configure(state=DISABLED)
    checkbox_divSul.configure(state=DISABLED)
    checkbox_divNorte.configure(state=DISABLED)
    # Validação Controladoria
elif abrir == '30':
    radio_recebimento.configure(state=DISABLED) 
    radio_embarques.configure(state=DISABLED)
    radio_entrega.configure(state=DISABLED)
    checkbox_divNorte.configure(state=DISABLED)
    checkbox_divSul.configure(state=DISABLED)
    checkbox_fortaleza.configure(state=DISABLED)
    #radio_fortaleza.configure(state=DISABLED)
    # Validação Div Sul
elif abrir == '40':
    radio_estatistico.configure(state=DISABLED)
    checkbox_fortaleza.configure(state=DISABLED)
    checkbox_divNorte.configure(state=DISABLED)
    #Todos
elif abrir == '99':
    radio_recebimento.configure(state=NORMAL)
    radio_embarques.configure(state=NORMAL)
    #radio_fortaleza.configure(state=NORMAL)
    radio_estatistico.configure(state=NORMAL)
else:
    messagebox.showwarning('Erro', 'Usuário não identificado'),janela.close()
    radio_recebimento.configure(state=DISABLED)
    radio_embarques.configure(state=DISABLED)
    #radio_fortaleza.configure(state=DISABLED)
    radio_estatistico.configure(state=DISABLED)
    radio_entrega.configure(state=DISABLED)
    checkbox_divNorte.configure(state=DISABLED)
    checkbox_divSul.configure(state=DISABLED)
    checkbox_fortaleza.configure(state=DISABLED)


janela.mainloop()






