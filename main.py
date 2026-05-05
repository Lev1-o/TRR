#main.py
def main():
    while True:
        print("\n🚀 SISTEMA PMI")
        print("1 - Criar proposta")
        print("2 - Gestão de contratos")
        print("3 - Gestão de manutenção")
        print("0 - Sair")

        opcao = input("Escolha: ")

        if opcao == "1":
            from CODIGOS.PROPOSTA.CONTRATO import executar
            executar()

        elif opcao == "2":
            from CODIGOS.GESTAO.GESTAO_CONTRATOS import menu
            menu()

        elif opcao == "3":
            from CODIGOS.GESTAO.GESTAO_FROTA import menu
            menu()

        elif opcao == "0":
            break

if __name__ == "__main__":
    main()
