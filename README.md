# xlsx_to_kicad_sym

Este repositório contém scripts em python que fazem a extração de parâmetros de componentes eletrônicos a partir de uma tabela xlsx e gera arquivos kidcad_sym utilizando templates

Abstract: Python script that extracts parameters from a list of electronic component descriptions and creates kicad_sym files from templates




## como funciona

1. O arquivo xmlx é aberto e lido
2. Os componentes são filtrados de acordo com seu código, cada tipo de componente possui um parser de descrição dedicado
3. Os parâmetros são extraidos e gera-se os arquivos intermediários, que auxiliam na depuração
    *  catalogo_<smd | pth>.xlsx - contém os parâmetros extraidos separados em clunas
    * <capacitors | resistors>_<smd | pth>.txt - contém o template preenchido dos componentes, ainda não aplicado ao template da biblioteca
4. Gera-se a bliblioteca de simbolos    
    * <capacitors | resistors>_<smd | pth>.kicad_sym - contém o arquivo de biblioteca gerado

**Atenção**
    Durante a extração de parâmetros podem haver casos de parâmetros identificados em duplicidade ou não encontrados. Esses casos são classificados como erros ou warnings e são apresentados no arquivo xlsx intermediário

## pacotes requeridos

### openpyxl
Esse projeto utiliza openpyxl para acessar e criar arquivos xlsx, para instalar basta utilizar o comando pip conforme exemplo a baixo

    $ pip install openpyxl