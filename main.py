
def split_name(full_name):
    full_name_splited = full_name.split()
    first_name = full_name_splited[0]
    delimiter = " "
    last_name = delimiter.join(full_name_splited[1:])
    return  {'first_name':first_name, 'last_name':last_name}

def clean_cpf(cpf):
    cpf_novo = str(cpf).replace(".",'').replace('-','').strip()
    return cpf_novo
    

def get_user_list_from_xlsx():
    import openpyxl
    wb = openpyxl.load_workbook("Planilha-insercao-em-massa-4-cia.xlsx")

    planilha_nome = wb.sheetnames[0]

    planilha = wb[planilha_nome]

    users_list = []


    for row in planilha.values:        
        if  not row[1]:
            continue





        user_dict = {
            'nome_completo': str(row[0]).upper(),            
            'rg':str(row[1]),
            'cpf':clean_cpf(row[2]),
            'id_funcional':str(row[3]),
            'email':str(row[4]).lower(),
            'tel':str(row[5]),
        }

        


        if user_dict.get('nome_completo') and user_dict.get('nome_completo') != "NOME COMPLETO" :
            user_dict |= split_name(user_dict['nome_completo'])   # Adicionando first_name e last_name ao dicionário com a função split_name            
            users_list.append(user_dict)
        
    return users_list







def script_for_add_user_in_ad(user_dict):    
    """dsadd user cn="WIVIANE BÁRBARA DA SILVA SILVEIRA",OU=COP,OU=SEPM,DC=SEPM,DC=rj,DC=gov,DC=br,DC=local -samid "a110575" -upn "110575@SEPM.rj.gov.br.local" -fn "WIVIANE" -ln "BÁRBARA DA SILVA SILVEIRA" -display "WIVIANE BÁRBARA DA SILVA SILVEIRA" -email "barbarasilveira2406@gmail.com" -pwd "Abc12345" -disabled no """
    line_script = f"""dsadd user cn="{user_dict['nome_completo']}",OU=COP,OU=SEPM,DC=SEPM,DC=rj,DC=gov,DC=br,DC=local -samid "a{user_dict['rg']}" -upn "{user_dict['rg']}@SEPM.rj.gov.br.local" -fn "{user_dict['first_name']}" -ln "{user_dict['last_name']}" -display "{user_dict['nome_completo']}" -email "{user_dict['email']}" -tel "{user_dict['tel']}" -desc "{user_dict['cpf']}" -pwd "Abc12345" -disabled no """
    print(line_script)









def main():
    
    users_list = get_user_list_from_xlsx()
    for user_dict in users_list:
        script_for_add_user_in_ad(user_dict)





if __name__ == "__main__":
    main()

