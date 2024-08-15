# **Script de Automação para Análise de Critérios Formais de Revistas Científicas**

Este script em Python foi desenvolvido para automatizar o processo de análise dos critérios formais para submissões a uma revista científica. O script avalia vários aspectos formais dos manuscritos para garantir que eles atendam às diretrizes de submissão da Revista de Estudos Jurídicos da Universidade Estadual Paulista - REJ UNESP.

## **Funcionalidades**

- **Verificação Automatizada de Critérios**: Verifica automaticamente a conformidade com critérios formais, como formatação, estilo de referência e estrutura do documento.

- **Geração de Relatório**: Gera um relatório resumido destacando as áreas que não estão em conformidade com as diretrizes da revista.

## **Requisitos**

- Python 3.11
- docx (python-docx)
- tkinter

Para instalar as bibliotecas necessárias, execute:

```bash
pip install python-docx
```

## **Uso**

Coloque os manuscritos que você deseja analisar em uma pasta individual e selecione o diretório.

Execute o script:

```bash
revisor_textos_cientificos.py
```

O script gerará um relatório para cada manuscrito na pasta `output`, detalhando a formatação do texto do documento tal como tamanho da fonte, alinhamento, espaçamento entre linhas, nome das seções/subseções e formatação das referências bibliográficas.

## **Contribuições**

Contribuições são bem-vindas! Por favor, envie um pull request ou abra uma issue para discutir quaisquer alterações.

## **Licença**

Este projeto é licenciado sob a Licença MIT - consulte o arquivo `LICENSE` para obter detalhes.
