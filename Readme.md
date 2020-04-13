#Get
O objetivo deste pequeno script é agilizar a tomada de dados de um espectofotômetro em
um ensaio de fotocatálise. Cada pasta informada ao script gera uma saída em Excel.
O script foi idealizado para o método de saída de um espectofotômetro 
Agilent Cary 60.

##Instalação
Não é necessário instalação.

##Uso
O scipt irá gerar um arquivo em Excel com os dados encontrados.
Na construção o script leva algumas premissas, logo é preciso garantir que elas
são verdadeiras para os dados do usuário. Entre elas:

- [ ] O nome dos arquivos é, sem aspas: "algumaidentificação_temponoreator"
- [ ] Todos os arquivos de saída do UV-Vis estão devidamente separados. Isso porque 
dependendo da configuração, duas medidas podem sair no mesmo arquivo
- [ ] Dentro do intervalo do comprimento de onda informado só existe uma medição

Recomendações:
- [ ] Uma pasta para cada ensaio
- [ ] Use intervalos onde: (Limite superior - Limite inferior) < 1

##Contribuição
O código é fornecido para que seja possível qualquer alteração ou aprimoramento. Se
o código fonte não chegou até você, me contate em luishqnery@gmail.com. O arquivo 
Requirement.txt contém os pacotes utilizados. Para instalá-los é possível passar:
```
pip install -r Requirements.txt
```