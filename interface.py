import flet 

def main(pagina):
    titulo = flet.Text("Pagina 1")
    botao_iniciar = flet.ElevatedButton("Iniciar Chat")
    janela = flet.AlertDialog()
    print(janela)
    
    pagina.add(titulo)
    pagina.add(botao_iniciar)

# executa o programa    
flet.app(main,view=flet.WEB_BROWSER)
