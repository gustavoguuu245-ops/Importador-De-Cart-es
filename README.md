# Importador-De-Cart-es
Sistema em Python (Pandas/Tkinter) para importar, tratar e conciliar planilhas financeiras da Stone. Automatiza a separação de taxas (rotativo/mensal) por CNPJ e tipo de cartão, gerando um CSV padronizado pronto para importação no sistema contábil.

Ferramenta desenvolvida para otimizar a rotina do setor financeiro e contábil, eliminando mêses de conferência manual de planilhas. O sistema lê relatórios brutos da máquina de cartão (Stone), mapeia automaticamente os CNPJs das filiais e aplica regras de negócio complexas para separar valores de crédito e débito, taxas de rotativo e mensalista oque levava dias para fazer agora e feito em praticamente 3 cliques. 

Ao final, o programa gera um arquivo CSV perfeitamente formatado e pronto para ser importado diretamente no sistema ERP/Contábil da empresa.

Principais Funcionalidades
Tratamento de Dados com Pandas: Leitura veloz de planilhas Excel (.xlsx) e limpeza de dados (remoção de espaços, tratamento de datas e valores).

Mapeamento de Regras de Negócio: Identificação inteligente de bandeiras (Visa, Master, Elo) e separação de códigos de pagamento.

Interface Gráfica Intuitiva: Tela desenvolvida em Tkinter para que o usuário final possa selecionar facilmente a empresa (CNPJ) que deseja processar.

Agrupamento e Consolidação: Soma os valores brutos por dia e por tipo de transação, garantindo que o financeiro bata centavo por centavo antes da exportação.

Tecnologias Utilizadas
Python: Linguagem principal.
Pandas: Para engenharia, filtragem e consolidação dos dados.
Tkinter: Para a interface de seleção e alertas de sucesso/erro.
