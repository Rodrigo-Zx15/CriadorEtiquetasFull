import tkinter as tk
from tkinter import filedialog
import xml.etree.ElementTree as ET
from tkinter import ttk

#o primeiro passo é capturar o filepath do xml - já nessa função, lemos o XML chamamos outra função para
#mostrar algumas informações da NF-e...
def select_xml():
    filePath = filedialog.askopenfilename(
        title="Selecione um arquivo xml.",
        filetypes=[("Arquivos xml de nota fiscal","*.xml"),("Todos os arquivos","*.*")]
    )
    if filePath:
        try:
            filePathLabel.config(text=f'Caminho do arquivo: {filePath}')
            showInfo(filePath)
            generateLabelsButton.grid()
        except Exception as e:
            label = ttk.Label(window,text=f"Erro ao gerar etiquetas: {e}")
            label.pack(pady=60)


#...e essa função preenche os textos da janela com as informações da NFe escolhida
def showInfo(xmlFile):
    tree = ET.parse(xmlFile)
    root = tree.getroot()
    nsNF = {'ns':"http://www.portalfiscal.inf.br/nfe"}
    nfData = {
        "emissao":'',
        "empresa":'',
        "itens":0
    }
    nfData["emissao"] = root.find("ns:NFe/ns:infNFe/ns:ide/ns:dhEmi",nsNF).text
    nfData["empresa"] = root.find("ns:NFe/ns:infNFe/ns:emit/ns:xNome",nsNF).text
    for item in root.findall('ns:NFe/ns:infNFe/ns:det/ns:prod/ns:qCom',nsNF):
        nfData["itens"] = nfData["itens"] + int(item.text)
    nfDateLabel.config(text=f'Data de emissão:{nfData["emissao"]}')
    nfCompanyLabel.config(text=f'Empresa:{nfData["empresa"]}')
    nfItemsLabel.config(text=f'Unidades:{nfData["itens"]}')

 #essa função lê a tabela de produtos do xml, mas a função não está parametrizada para qualquer nfe -
 #a chave "codigo" do dict "produto" só capturará o SKU único que vem nas nfes da Amazon,
 #por mais que tecnicamente a função consiga ler qualquer NFe.

def get_products():
    xmlFile = filePathLabel.cget("text").split(sep='arquivo:')[1].strip()
    tree = ET.parse(xmlFile)
    root = tree.getroot()
    nsNF = {'ns': "http://www.portalfiscal.inf.br/nfe"}
    products = []
    for item in root.findall('ns:NFe/ns:infNFe/ns:det',nsNF):
        product = {
            "sku":'',
            "titulo":'',
            "quantidade": 0,
            "codigo":''
        }
        
        product['sku'] = item.find('ns:prod/ns:cProd',nsNF).text
        product['titulo'] = item.find('ns:prod/ns:xProd',nsNF).text.split(',')[0]
        product['quantidade'] = int(item.find('ns:prod/ns:qCom',nsNF).text)
        product['codigo'] = item.find('ns:infAdProd',nsNF).text.split(':')[1].replace(';','')
        products.append(product)
    #return products
    labelColumns = labelOptionsBox.get()
    generate_labels(products, labelColumns)
    

#aqui, geramos o código ZPL que modela as etiquetas das impressoras zebras

def generate_labels(products,labelColumns):
    with open("./etiquetas_amazon.txt","w") as file:
        file.write("")
    model = ''
    labelColumns = labelOptionsBox.get()
    #para rolos de etiquetas de 1 coluna, basta "empilhar" um bloco de código ZPL por produto após o outro
    #se quiser entender mais o que cada linha da variável model significa, sugiro usar o website laberaly:
    #https://labelary.com/zpl.html
    if labelColumns == '1 coluna':
        for product in products:
       
            model = f"""
            ^XA
            ~TA000
            ~JSN
            ^LT0
            ^MNW
            ^MTT
            ^PON
            ^PMN
            ^LH0,0
            ^JMA
            ^PR8,8
            ~SD23
            ^JUS
            ^LRN
            ^CI27
            ^PA0,1,1,0
            ^XZ
            ^XA
            ^MMT
            ^PW320
            ^LL200
            ^LS0
            ^FT7,24^A0N,18,23^FH\^CI28^FDAMAZON^FS^CI27
            ^BY2,3,70^FT15,100^BCN,,Y,N
            ^FH\^FD>:{product["codigo"]}^FS
            ^FT7,158^A0N,15,15^FB345,2,1^FH\^CI28^FD{product["titulo"]}^FS^CI27
            ^FT7,185^A0N,14,15^FH\^CI28^FDSKU:^FS^CI27
            ^FT41,185^A0N,20,15^FH\^CI28^FD{product["sku"]}^FS^CI27
            ^PQ{product["quantidade"]},0,1,Y
            ^XZ\n
            """
            with open("./etiquetas_amazon.txt","a") as file:
                file.write(model)
    

#a lógica para geração de etiquetas para rolos de 2 colunas é mais complicada, mas simplificadamente:
#resolvi trabalhar ao redor da divisão por 2 - o programa gerará x linhas de pares de etiquetas,
#onde x = unidades do produto/2. Caso unidades seja um nº ímpar, uma etiqueta solitária é inserida logo em seguida.
    if labelColumns == "2 colunas":
        with open("./etiquetas_amazon.txt","a") as file:
            file.write("^XA" \
            "^PW640"
            "^MMT"
            "^LL200"
            "^SD23" \
            "^XZ")
        for product in products:
            if product["quantidade"] % 2 == 0:
                model = f"""
                    ^XA
            ^FT15,24^A0N,18,23^FH\^CI28^FDAMAZON^FS^CI27
            ^BY2,3,70^FT15,100^BCN,,Y,N
            ^FH\^FD>:{product["codigo"]}^FS
            ^FT15,158^A0N,15,15^FB335,2,1^FH\^CI28^FD{product["titulo"]}^FS^CI27
            ^FT15,185^A0N,14,15^FH\^CI28^FDSKU:^FS^CI27
            ^FT50,185^A0N,20,15^FH\^CI28^FD{product["sku"]}^FS^CI27
            
            ^FT345,24^A0N,18,23^FH\^CI28^FDAMAZON^FS^CI27
            ^BY2,3,70^FT340,100^BCN,,Y,N
            ^FH\^FD>:{product["codigo"]}^FS
            ^FT340,158^A0N,15,15^FB345,2,1^FH\^CI28^FD{product["titulo"]}^FS^CI27
            ^FT340,185^A0N,14,15^FH\^CI28^FDSKU:^FS^CI27
            ^FT370,185^A0N,20,15^FH\^CI28^FD{product["sku"]}^FS^CI27
            ^PQ{int(product["quantidade"]/2)},0,1,Y
                    ^XZ
                """
                with open("./etiquetas_amazon.txt","a") as file:
                    file.write(model)
            elif product["quantidade"] == 1:
                model = f"""
                    ^XA
            ^FT15,24^A0N,18,23^FH\^CI28^FDAMAZON^FS^CI27
            ^BY2,3,70^FT15,100^BCN,,Y,N
            ^FH\^FD>:{product["codigo"]}^FS
            ^FT10,158^A0N,15,15^FB345,2,1^FH\^CI28^FD{product["titulo"]}^FS^CI27
            ^FT10,185^A0N,14,15^FH\^CI28^FDSKU:^FS^CI27
            ^FT41,185^A0N,20,15^FH\^CI28^FD{product["sku"]}^FS^CI27
            ^PQ1,0,1,Y
                ^XZ
                """
                with open("./etiquetas_amazon.txt","a") as file:
                    file.write(model)
            else:
                model = f"""
                    ^XA
            ^FT15,24^A0N,18,23^FH\^CI28^FDAMAZON^FS^CI27
            ^BY2,3,70^FT15,100^BCN,,Y,N
            ^FH\^FD>:{product["codigo"]}^FS
            ^FT10,158^A0N,15,15^FB345,2,1^FH\^CI28^FD{product["titulo"]}^FS^CI27
            ^FT10,185^A0N,14,15^FH\^CI28^FDSKU:^FS^CI27
            ^FT41,185^A0N,20,15^FH\^CI28^FD{product["sku"]}^FS^CI27
            
            ^FT345,24^A0N,18,23^FH\^CI28^FDAMAZON^FS^CI27
            ^BY2,3,70^FT340,100^BCN,,Y,N
            ^FH\^FD>:{product["codigo"]}^FS
            ^FT340,158^A0N,15,15^FB345,2,1^FH\^CI28^FD{product["titulo"]}^FS^CI27
            ^FT340,185^A0N,14,15^FH\^CI28^FDSKU:^FS^CI27
            ^FT370,185^A0N,20,15^FH\^CI28^FD{product["sku"]}^FS^CI27
            ^PQ{int(product["quantidade"]/2)},0,1,Y
                    ^XZ
                """
                with open("./etiquetas_amazon.txt","a") as file:
                    file.write(model)
                model = f"""
                        ^XA
                ^FT15,24^A0N,18,23^FH\^CI28^FDAMAZON^FS^CI27
                ^BY2,3,70^FT15,100^BCN,,Y,N
                ^FH\^FD>:{product["codigo"]}^FS
                ^FT7,158^A0N,15,15^FB345,2,1^FH\^CI28^FD{product["titulo"]}^FS^CI27
                ^FT7,185^A0N,14,15^FH\^CI28^FDSKU:^FS^CI27
                ^FT41,185^A0N,20,15^FH\^CI28^FD{product["sku"]}^FS^CI27
                ^PQ1,0,1,Y
                    ^XZ
                    """
                with open("./etiquetas_amazon.txt","a") as file:
                        file.write(model)
    show_contact_info()
#gerando a tela inicial, com o grid 
window = tk.Tk()
window.title("Criador de etiquetas")
mainframe = ttk.Frame(window)
mainframe.grid(column=0,row=0,sticky=['N','W','E','S'])
window.columnconfigure(0,weight=1)
window.rowconfigure(0,weight=1)

#criando labels vazias
filePathLabel = ttk.Label(window,text='Caminho do arquivo:-----')
filePathLabel.grid(column=1,row=0,sticky='W')
nfDateLabel = ttk.Label(window,text='Data de Emissão:-----')
nfDateLabel.grid(column=1,row=1,sticky='W')
nfCompanyLabel = ttk.Label(window,text='Empresa:-----')
nfCompanyLabel.grid(column=1,row=2,sticky='W')
nfItemsLabel = ttk.Label(window,text='Unidades:-----')
nfItemsLabel.grid(column=1,row=3,sticky='W')

#criando lista suspensa p/ selecionar quantas colunas tem o rolo de etiquetas
labelOptions = ['1 coluna','2 colunas']
labelOptionsBox = ttk.Combobox(window,values=labelOptions)
labelOptionsBox.grid(column=2, row=4)
labelOptionsBox.set("Colunas?")

showInfoButton = ttk.Button(window,text="Selecione a NFe",command=select_xml)
showInfoButton.grid(column=1,row=4,sticky=['S','N'])


generateLabelsButton = ttk.Button(window,text="Gerar etiquetas",command=get_products)
generateLabelsButton.grid(column=0,row=4,sticky=['S','N'])
generateLabelsButton.grid_remove()


#autopromoção, pois por que não? :)
def show_contact_info():
    contato = tk.Toplevel()
    contato.title("Informações de Contato")
    contato.geometry("300x180")
    contato.resizable(False, False)

    ttk.Label(contato, text="📬 Suas etiquetas estão prontas! \nSe gostou, entre em contato e confira outros projetos:", font=("Arial", 9)).pack(pady=10)

    ttk.Label(contato, text="Criado por: Rodrigo Nascimento Gomes").pack()
    ttk.Label(contato, text="Email: rogui.2000@hotmail.com").pack()
    ttk.Label(contato, text="LinkedIn: linkedin.com/in/rodrigo-n-gomes/").pack()
    ttk.Label(contato, text="GitHub: github.com/Rodrigo-Zx15").pack()

    closeButton = ttk.Button(contato, text="Fechar",command=contato.destroy)
    closeButton.pack(pady=15)
    

window.mainloop()
#generateLabelsButton = tk.Button(window,text="Gerar etiquetas!",command=xmlReader.generate_labels)