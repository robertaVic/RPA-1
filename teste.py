def recuperar_diretorio():
    diretorio = os.path.realpath(__file__)
    diretorio = diretorio.split("\\")
    auxDir = ""
    auxDir1 = ""
    for d in range(len(diretorio)-1):
        auxDir+=diretorio[d]
        auxDir+="\\"
        if d != len(diretorio)-2:
            auxDir1+=diretorio[d]
            auxDir1+="\\"
    diretorioA = auxDir
    return diretorioA

print(str(recuperar_diretorio))