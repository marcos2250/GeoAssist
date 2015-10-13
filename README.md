GeoAssist
=========

GeoAssist é uma ferramenta simples e útil para trabalhos de Mapeamento e Georeferenciamento de Imóveis Rurais.

O GeoAssist oferece as seguintes funções:

* Ambiente CAD simplificado para desenho de poligonais 3D;
* Visualização 2D precisa, fácil de usar;
* Cadastramento de confrontantes, com visualização no mapa;
* Cadastramento das informações imobiliárias e pessoais;
* Cálculo de área, perímetro, azimutes e cotas;
* Cálculo de coordenadas geodésicas;
* Cálculo da convergência meridiana e fator de escala;
* Conversão de coordenadas cartesianas em polares, e vice-versa;
* Conversões de coordenadas entre Datums, UTM e Latitude/Longitude;
* Emissão de Memoriais Descritivos e Planilhas de Coordenadas;
* Emissão de Relatórios para Georeferenciamento;
* Compatibilidade com o Google Earth, para visualização do terreno;
* Produção automatizada de plantas em formato DXF (AutoCAD);
* Impressão de croquis diretamente da tela do programa.

O GeoAssist possui as seguintes limitações:
* A poligonal inteira deve estar contida em uma única zona UTM, o
  GeoAssist ainda não permite trabalhar com áreas divididas por 2 ou
  mais zonas UTM;
* Havendo ilhas internas no imóvel, as áreas devem ser trabalhadas
  separadamente, e suas poligonais agrupadas manualmente, por via
  de outro software CAD externo (AutoCAD e outros);



###Plataforma e Requerimentos

* Desenvolvido para a plataforma Windows 32 e 64 bits, 2000/XP/2007 etc.
* Linguagem/IDE utilizada: Visual Basic 6.
* Recomendável Office 2003 (com service packs), 2007 ou superior instalado. 
* Opcionalmente, é recomendável também o Google Earth desktop e AutoCAD 2000 ou superior.



###Compilando o projeto e instalação

1. Para montar o ambiente de desenvolvimento, é necessário possuir o Visual Basic 6.
2. O módulo SuperFlexGrid (src/superFlexGrid/SuperGrid.vbp) é um componente OCX utilizado pelo GeoAssist, é necessário primeiramente compilar o arquivo SuperGrid.OCX e registrá-lo no ambiente do Windows (System32/SYSWOW64).
2. Em seguida, importar o projeto principal (src/GeoAssist.vbp), verificar se as dependências estão ok e compilar seu executável GeoAssist.EXE.
3. Para "deploy", teste ou debug, juntar os 2 artefatos gerados, mais os arquivos contidos em *resources* no mesmo diretório. 

Para instalação manual em outros PCs, basta copiar os arquivos compilados e os arquivos contidos em *resources*, e executar o **install.bat**.
O WinRAR pode ser usado para montar um instalador automático (SFX), bastando compactar os arquivos da pasta em modo SFX, e acrescentando o conteúdo do script *installsfxdata.txt* na aba "Comentários".
