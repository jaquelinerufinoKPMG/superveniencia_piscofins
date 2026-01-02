import pandas as pd
import locale

class DocumentFormatter:
    def __init__(self):
        pass
        
    def format_documents(cpf_cnpj: str) -> str:
        if pd.isnull(cpf_cnpj):
            return cpf_cnpj
        cpf_cnpj = str(cpf_cnpj).strip()
        cpf_cnpj = "".join(filter(str.isdigit, cpf_cnpj))  # Mantém apenas números

        if len(cpf_cnpj) == 14:  # CNPJ
            return f"{cpf_cnpj[:2]}.{cpf_cnpj[2:5]}.{cpf_cnpj[5:8]}/{cpf_cnpj[8:12]}-{cpf_cnpj[12:]}"
        elif len(cpf_cnpj) == 11:  # CPF
            return f"{cpf_cnpj[:3]}.{cpf_cnpj[3:6]}.{cpf_cnpj[6:9]}-{cpf_cnpj[9:]}"
        elif len(cpf_cnpj) < 11:  # CPF incompleto, preencher com zeros
            cpf_cnpj = cpf_cnpj.zfill(11)
            return f"{cpf_cnpj[:3]}.{cpf_cnpj[3:6]}.{cpf_cnpj[6:9]}-{cpf_cnpj[9:]}"
        elif len(cpf_cnpj) < 14:  # CNPJ incompleto, preencher com zeros
            cpf_cnpj = cpf_cnpj.zfill(14)
            return f"{cpf_cnpj[:2]}.{cpf_cnpj[2:5]}.{cpf_cnpj[5:8]}/{cpf_cnpj[8:12]}-{cpf_cnpj[12:]}"

        return cpf_cnpj  # Retorna o valor original se não for possível formatar


    def format_date_columns(df: pd.DataFrame, date_columns: list[str]):
        for col in date_columns:
            if col in df.columns:
                # Converte em bloco, assumindo formato DD/MM/YYYY, força coerção de erros
                df[col] = (
                    pd.to_datetime(
                        df[col], dayfirst=True, format="%d/%m/%Y", errors="coerce"
                    )
                    .dt.strftime("%d/%m/%Y")
                    .fillna("")
                )
            else:
                print(f"Warning: Column '{col}' does not exist in the DataFrame.")
        return df


    def format_values(amount, format_as_currency=False):

        if pd.isnull(amount):
            return amount
        amount = str(amount)
        try:
            # Configura o locale para o Brasil
            locale.setlocale(locale.LC_ALL, "pt_BR.UTF-8")

            # Remove pontos e converte vírgulas em pontos para valores com separador brasileiro
            if "," in amount and "." in amount:
                amount = amount.replace(".", "").replace(",", ".")
            elif "," in amount:
                amount = amount.replace(",", ".")

            # Converte para float
            try:
                amount_as_float = float(amount)
            except ValueError:
                raise ValueError(f"Não foi possível converter o valor: {amount}")

            # Formata o número no padrão monetário brasileiro
            formatted_value = locale.currency(
                abs(amount_as_float), grouping=True, symbol=True
            )

            # Adiciona o sinal negativo para números negativos
            if amount_as_float < 0:
                formatted_value = formatted_value.replace("R$", "- R$")

            return formatted_value if format_as_currency else amount_as_float
        except Exception as e:
            # print(f"Erro ao formatar o valor '{valor}': {e}")
            return amount


    def to_pascal_case(text_input: str) -> str:
        if pd.isnull(text_input):
            return text_input
        return " ".join(word.capitalize() for word in text_input.split(" "))
    
    def correct_year(data):
        if pd.isnull(data):
            return data
        try:
            partes = data.split('/')
            if len(partes) == 3:
                dia, mes, ano = partes
                if len(ano) == 2:
                    if int(ano) > 40:
                        ano_corrigido = '19' + ano
                    else:
                        ano_corrigido = '20' + ano
                    return f"{dia}/{mes}/{ano_corrigido}"
            return data
        except Exception:
            return data
