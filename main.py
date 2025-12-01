
def split_name(full_name):
    full_name_splited = full_name.split()
    first_name = full_name_splited[0]
    delimiter = " "
    last_name = delimiter.join(full_name_splited[1:])
    return  {'first_name':first_name, 'last_name':last_name}



def get_user_list_from_xlsx():
    import openpyxl
    wb = openpyxl.load_workbook("users-ad.xlsx")

    planilha_nome = wb.sheetnames[0]

    planilha = wb[planilha_nome]

    users_list = []


    for row in planilha.values:
        # print(row)


        user_dict = {
            'nome_completo': row[1],            
            'rg':str(row[2]),
            'cpf':row[3],
            'id_funcional':str(row[4]),
            'email':row[5],
            'tel':str(row[6]),
        }


        if user_dict.get('nome_completo') and user_dict.get('nome_completo') != "NOME COMPLETO" :
            user_dict |= split_name(user_dict['nome_completo'])   # Adicionando first_name e last_name ao dicionário com a função split_name            
            users_list.append(user_dict)
        
    return users_list







def script_for_add_user_in_ad(user_dict):    
    """dsadd user cn="WIVIANE BÁRBARA DA SILVA SILVEIRA",OU=COP,OU=SEPM,DC=SEPM,DC=rj,DC=gov,DC=br,DC=local -samid "a110575" -upn "110575@SEPM.rj.gov.br.local" -fn "WIVIANE" -ln "BÁRBARA DA SILVA SILVEIRA" -display "WIVIANE BÁRBARA DA SILVA SILVEIRA" -email "barbarasilveira2406@gmail.com" -pwd "Abc12345" -disabled no """
    line_script = f"""dsadd user cn="{user_dict['nome_completo']}",OU=COP,OU=SEPM,DC=SEPM,DC=rj,DC=gov,DC=br,DC=local -samid "a{user_dict['rg']}" -upn "{user_dict['rg']}@SEPM.rj.gov.br.local" -fn "{user_dict['first_name']}" -ln "{user_dict['last_name']}" -display "{user_dict['nome_completo']}" -email "{user_dict['email']}" -pwd "Abc12345" -disabled no """
    print(line_script)









def main():
    print("Imprimindo lista de usuários:")
    users_list = get_user_list_from_xlsx()
    for user_dict in users_list:
        script_for_add_user_in_ad(user_dict)





if __name__ == "__main__":
    main()

