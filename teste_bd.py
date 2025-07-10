import pandas as pd
from app import app, db, Materiais

def to_int_safe(value):
    try:
        return int(value)
    except (ValueError, TypeError):
        return 0

def importar_materiais():
    # Lê o arquivo Excel (coloque seu caminho correto)
    df = pd.read_excel(r'C:\Users\Eduardo Prado\OneDrive\Desktop\estoque.xlsx')

    # Remove espaços extras nos nomes das colunas
    df.columns = df.columns.str.strip()

    # Imprime as colunas para conferência
    print("Colunas no Excel:", df.columns.tolist())

    empresa_padrao = "Empresa Padrão"
    aplicacao_padrao = "Aplicação não especificada"

    with app.app_context():
        for _, row in df.iterrows():
            QuantidadeEstoque = to_int_safe(row['QuantidadeEstoque'])

            material = Materiais(
                DescricaoMaterial=row['Material'],
                Empresa=empresa_padrao,
                Aplicacao=aplicacao_padrao,
                QuantidadeEstoque=QuantidadeEstoque,
                Fornecedor=None if row['Fornecedor'] == '-' else row['Fornecedor'],
                NumeroNF=None if row['NumeroNF'] == '-' else row['NumeroNF'],
                FatorConsumo=0.0,
                Ativo=False
            )
            db.session.add(material)
        db.session.commit()

    print("Importação finalizada!")

if __name__ == "__main__":
    importar_materiais()
