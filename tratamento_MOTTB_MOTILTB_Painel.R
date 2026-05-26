################################################################################################################################
################################################       20 DE AGOSTO DE 2024        #############################################
###############################################       PAINEL ILTB, MOT-ILTB e MOT_TB   #################################################
##############################################            versão R 4.0.3            ############################################
############################################            CASSIA CP MENDICINO         ############################################
########################################### MY LORDY GIVE-ME STRAING TO CARRY ON    ############################################
################################## SANCTA MARIA, MATER DEI, ORA PRO NOBIS PECCATORIBUS #########################################
################################################################################################################################


#FONTE DE INFORMAÇÕES: 
#RELATORIO SISTEMA INFORMAÇAO DE MONITORAMENTO CLÍNICO (SIMC) NOMINAIS a partir de janeiro de  2023  SEM FILTRO
    #Necessário abrir o banco bruto SIMC, excluir o cabeçalho e salvar como Excel Workbook
#RELATORIO SISTEMA DE INFECÇÃO LATENTE PARA TUBERCULOSE (IL-TB) NOMINAIS E SEM FILTRO
#RELATORIO SISTEMA DE NOTIFICAÇÃO DE AGRAVOS DE NOTIFICAÇÕES PARA TUBERCULOSE DE MINAS GERAIS A PARTIR DO ANO DE 2000
#TABELAS DE MUNICÍPIOS PDR 2024 

#link do painel:https://app.powerbi.com/view?r=eyJrIjoiMGIzMjNlMDEtZDViMy00MmVkLWE0NzgtZTQ0ZmNlNjQ2Y2IzIiwidCI6Ijg3ZTRkYTJiLTgyZGYtNDhmNi05MTU3LTY5YzNjYTYwMGRmMiIsImMiOjR9

rm(list = ls())
library(writexl)#Exportar em Excel
library(data.table)
library(tidyr)# função pivot (emparelhamento de colunas)
library(dplyr) 
library(openxlsx)
library(readxl)
library(foreign)
library(lubridate)
library(dplyr) #função mutate
library(readr)#EXPORTAR EM CSV


#BANCO SIMC: MARCAR AÇÃO RECOMENDADA: TODAS E PACIENTES: TODOS

#### IMPORTAÇÃO 
SIMC_bruto<-read.xlsx("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/SIMC_bruto_2025_atual.xlsx") #57831
ILTB_bruto<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/banco_bruto_ILTB.xlsx") #16359
MUNICIPIOS<-read.xlsx("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/TABELA DE MUNICÍPIOS PDR 2026.xlsx") 
SINAN_TB_bruto<-read.dbf("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/TUBENET.dbf") #124033
MUNICIPIOS_POP<-read.xlsx("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/PDR2026_comPOP.xlsx") 
USUARIOS_BRUTO<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Usuários_ILTB.xlsx")#3207

#SES





####################################################   ILTB tratamento geral   #####################################################

# Seleção das variáveis:
ILTB<-ILTB_bruto[,c("Data de início do tratamento atual","Data de nascimento","Data de notificação","Data do término do tratamento",
                    "Data do encerramento","Situação de Encerramento","Nome de registro","Medicamentos","Gestante","Indicação de Tratamento",
                    "Sexo","Unidade de tratamento inicial - Município (Cod. IBGE)","Número do caso",
                    "Dígito do caso","Tipo de entrada")]


####    TRATAMENTO E CRIAÇÃO DE VARIÁVEIS   ####

names(ILTB)[names(ILTB)=="Data de início do tratamento atual"]<-"DT_TT"
names(ILTB)[names(ILTB)=="Data de nascimento"]<-"DT_NASC"
names(ILTB)[names(ILTB)=="Data de notificação"]<-"DT_NOTIF"
names(ILTB)[names(ILTB)=="Data do término do tratamento"]<-"DT_TT_final"
names(ILTB)[names(ILTB)=="Data do encerramento"]<-"DT_ENCERRAMENTO"
names(ILTB)[names(ILTB)=="Situação de Encerramento"]<-"ENCERRAMENTO"
names(ILTB)[names(ILTB)=="Nome de registro"]<-"Paciente"
names(ILTB)[names(ILTB)=="Indicação de Tratamento"]<-"INDIC_TT"

ILTB$DT_TT<- as.Date(ILTB$DT_TT, origin = "1899-12-30")
ILTB$DT_NASC<- as.Date(ILTB$DT_NASC, origin = "1899-12-30")
ILTB$DT_NOTIF<- as.Date(ILTB$DT_NOTIF, origin = "1899-12-30")
ILTB$DT_TT_final<- as.Date(ILTB$DT_TT_final, origin = "1899-12-30")

    #Criar variável "Tratamento" finalizado/não finalizado a partir dos campos vazios na variável encerramento
ILTB <- ILTB %>%
  mutate(
    Tratamento = if_else(is.na(ENCERRAMENTO), "Não finalizado", "Finalizado")
  )

#Calcular o tempo de tratamento: casos não encerrados: data de tratamento até o dia atual, casos encerrados: 
#data do início do tratamento menos data final, em meses:
ILTB <- ILTB %>%
  mutate(
    Tempo_TT = case_when(
      Tratamento == "Não finalizado" ~ round(as.numeric(Sys.Date() - DT_TT) / 30.44),
      Tratamento == "Finalizado" ~ round(as.numeric(DT_TT_final - DT_TT) / 30.44),)
  )

#Variável TT_EXPIRADO:Tempo de acordo com o tratamento:
ILTB <- ILTB %>% 
  mutate(
    TT_EXPIRADO = case_when(
      Medicamentos == "Isoniazida" & Tempo_TT > 12 ~ "Expirado",
      Medicamentos == "Isoniazida - 6H" & Tempo_TT > 6 ~ "Expirado",
      Medicamentos == "Isoniazida - 9H" & Tempo_TT > 9 ~ "Expirado",
      Medicamentos == "Rifampicina - 4R" & Tempo_TT > 4 ~ "Expirado",
      Medicamentos == "Rifampicina + Isoniazida - 3RH (dispersíveis pediátricos)" & Tempo_TT > 3 ~ "Expirado",
      Medicamentos == "Rifapentina + Isoniazida - 3HP" & Tempo_TT > 3 ~ "Expirado",
      TRUE ~ "Não Expirado"
    )
  )

#Variavel ANO:
ILTB$ano<-format(ILTB$DT_TT,"%Y")

#Variável IDADE e IDADE_cat;
ILTB$IDADE<-round((as.numeric(ILTB$DT_TT-ILTB$DT_NASC))/365.25)
#IDADE_CAT
ILTB$IDADEcat<-rep(NA,length(ILTB$IDADE))
ILTB$IDADEcat[ILTB$IDADE<2]<-"Inferior a 2 anos" 
ILTB$IDADEcat[2<=ILTB$IDADE & ILTB$IDADE<10]<-"3 a 9 anos"
ILTB$IDADEcat[10<=ILTB$IDADE & ILTB$IDADE<15]<-"10 a 14 anos"
ILTB$IDADEcat[15<=ILTB$IDADE & ILTB$IDADE<20]<-"15 a 19 anos"
ILTB$IDADEcat[20<=ILTB$IDADE & ILTB$IDADE<30]<-"20 a 29 anos"
ILTB$IDADEcat[30<=ILTB$IDADE & ILTB$IDADE<40]<-"30 a 39 anos"
ILTB$IDADEcat[40<=ILTB$IDADE & ILTB$IDADE<50]<-"40 a 49 anos"
ILTB$IDADEcat[50<=ILTB$IDADE & ILTB$IDADE<60]<-"50 a 59 anos"
ILTB$IDADEcat[ILTB$IDADE>=60]<-"Superior a 60 anos"
ILTB$IDADEcat=as.factor(ILTB$IDADEcat)


#Ordem idade: para ordenar no BI:
ILTB$IDADEordem<-rep(NA,length(ILTB$IDADEcat))
ILTB$IDADEordem[ILTB$IDADEcat=="Inferior a 2 anos"]<-9 
ILTB$IDADEordem[ILTB$IDADEcat=="3 a 9 anos"]<-8
ILTB$IDADEordem[ILTB$IDADEcat=="10 a 14 anos"]<-7
ILTB$IDADEordem[ILTB$IDADEcat=="15 a 19 anos"]<-6
ILTB$IDADEordem[ILTB$IDADEcat=="20 a 29 anos"]<-5
ILTB$IDADEordem[ILTB$IDADEcat=="30 a 39 anos"]<-4
ILTB$IDADEordem[ILTB$IDADEcat=="40 a 49 anos"]<-3
ILTB$IDADEordem[ILTB$IDADEcat=="50 a 59 anos"]<-2
ILTB$IDADEordem[ILTB$IDADEcat=="Superior a 60 anos"]<-1
ILTB$IDADEordem=as.numeric(ILTB$IDADEordem)


#VARIÁVEL GESTANTE  #################  FALTA COLOCAR ESTA INFORMAÇÃO NO PAINEL: CHAMAR ATENÇÃO
# PARA A CATEGORIA "nÃO SABE"
#Criar a categoria "Não se aplica"
#NÃO SE APLICA: Homens e mulheres abaixo de 10 anos e mulheres acima de 60 anos
ILTB$Gestante <- ifelse(
  ILTB$Sexo == "Masculino" | ILTB$IDADE < 10 | ILTB$IDADE > 60, 
  "NÃO SE APLICA", 
  ILTB$Gestante
)
#Juntar as categorias Ignorado e Não sabe:
ILTB$Gestante <- gsub("Ignorado","Não sabe",ILTB$Gestante)

#Variável Situação.de.encerramento
ILTB<-ILTB%>%
  mutate(
    TT_completo = case_when(
      ENCERRAMENTO %in% c("Interrupção do tratamento","Suspenso por condição clínica desfavorável ao tratamento",
                          "Suspenso por reação adversa","Suspenso por PT < 5mm em quimioprofilaxia primária") ~ "Tratamento incompleto",
      ENCERRAMENTO %in% c("Tratamento completo","Óbito","Transferido para outro país","Tuberculose ativa") ~"Tratamento completo",
      TRUE ~ "Tratamento não encerrado"
    )
  )
                    
# Variável indicação de tratamento: reduzir  número de categorias
ILTB <- ILTB %>% 
  mutate(
    INDIC_TT_cat = case_when(
      INDIC_TT %in% c("Contatos de TB pulmonar ou laríngea confirmada por critério laboratorial",
                      "Contatos de TB pulmonar ou laríngea, adultos e crianças, independentemente da vacinação  prévia com BCG") ~ "Contatos TB pulmonar",
      INDIC_TT %in% c("Diabetes mellitus", "Insuficiência renal em diálise","Neoplasias de cabeça e pescoço, linfomas e outras neoplasias hematológicas",
                      "Neoplasias em terapia imunossupressora") ~ "Cond.crônicas/Neoplasia",
      INDIC_TT %in% c("Pessoas com alterações radiológicas fibróticas sugestivas de sequela de TB",
                      "Pessoas com calcificação isolada (sem fibrose) na radiografia") ~ "Alterações radiológicas",
      INDIC_TT %in% c("Pessoas que farão uso ou estão em uso de  imunobiológicos e/ou imunossupressores, incluindo corticosteroides (correspondente a >15mg de prednisona por mais de um mês)",
                      "Pessoas que farão uso ou estão em uso de imunobiológicos e/ou imunossupressores, incluindo corticosteroides (correspondente a >15mg de prednisona por mais de um mês) com radiografia de tórax com cicatriz radiológica de TB, sem tratamento anterior para TB.",
                      "Pessoas que farão uso ou estão em uso de imunobiológicos e/ou imunossupressores, incluindo corticosteroides (correspondente a >15mg de prednisona por mais de um mês) com registro documental de ter tido PT maior ou igual 5mm ou IGRA positivo e não submetido ao tratamento da ILTB na ocasião.",
                      "Pessoas que farão uso ou estão em uso de imunobiológicos e/ou imunossupressores, incluindo corticosteroides (correspondente a >15mg de prednisona por mais de um mês) contatos de TB pulmonar ou laríngea com confirmação laboratorial.",
                      "Pessoas candidatas a transplante de células-tronco e/ou órgãos sólidos") ~ "Uso de Imunossupressores",
      INDIC_TT %in% c("Pessoas tabagistas (> 1 maço/dia)", "Silicose", "Pessoas com baixo peso (< 85% do peso ideal)") ~ "Fumo/Silicose/baixo peso",
      INDIC_TT %in% c("Pessoas vivendo com HIV/aids com CD4+ maior que 350 cél/mm3",
                      "Pessoas vivendo com HIV/aids com contagem de células CD4+ menor ou igual a 350 cél/mm3",
                      "Pessoas vivendo com HIV/aids com radiografia de tórax com cicatriz radiológica de TB, sem tratamento anterior para TB.",
                      "Pessoas vivendo com HIV/aids com registro documental de ter tido PT  maior ou igual 5mm ou IGRA positivo e não submetido ao tratamento da ILTB na ocasião",
                      "Pessoas vivendo com HIV/aids contatos de TB pulmonar ou laríngea com confirmação laboratorial") ~ "Pessoas vivendo com HIV",
      INDIC_TT %in% c("Profissionais de saúde", "Trabalhadores de instituições de longa permanência") ~ "Profissionais de saúde/ILP",
      INDIC_TT %in% c("Recém-nascidos coabitantes de caso fonte de tuberculose (TB) pulmonar ou laríngea confirmado por critério laboratorial") ~ "Recém-nascidos",
      INDIC_TT %in% c("Outra") ~ "Outras condições",
      TRUE ~ "INDETERMINADO"
    )
  )
table(ILTB$INDIC_TT_cat) #Verificar se tem alguém indeterminado


#"Unidade Regional de Saúde"  
MUNICIPIOS$Unidade.Regional.de.Saúde<-gsub("Diamantina", "DIAMANTINA",MUNICIPIOS$Unidade.Regional.de.Saúde)
MUNICIPIOS_ILTB<-MUNICIPIOS[,c("CÓDIGO","Unidade.Regional.de.Saúde","NM_MUNICIP")]
names(MUNICIPIOS_ILTB)[names(MUNICIPIOS_ILTB)=="NM_MUNICIP"]<- "MUNIC_INSTITUICAO"
names(ILTB)[names(ILTB)=="Unidade de tratamento inicial - Município (Cod. IBGE)"]<- "CÓDIGO"
ILTB<- merge(ILTB, MUNICIPIOS_ILTB, by = c("CÓDIGO")) 
# Linhas excluídas correspondem aos municípios de outros estados


## EXportação do banco para alimentação do painel ILTB nas páginas "Perfil da população" e "Quantitativos" 14.363 obs e 25 variables
#write.csv(ILTB,"C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/IL-TB/Painel/ILTB_TT_perfil.csv") 
#write.csv(ILTB,"C:/Users/x3996631/OneDrive - Secretaria de Estado de Saude de Minas Gerais/IL-TB/Painel/ILTB_TT_perfil.csv")



####################################################  Slide 1: Distribuição de casos no estado    ####################################################

ILTB_mapas<-ILTB[,c("CÓDIGO","Unidade.Regional.de.Saúde","MUNIC_INSTITUICAO","ano","Número do caso")]

#Mudar alguns nomes de regionais e de municípios para localização no mapa:
ILTB_mapas$Unidade.Regional.de.Saúde<-gsub("GOV. VALADARES","Governador Valadares",ILTB_mapas$Unidade.Regional.de.Saúde)
ILTB_mapas$Unidade.Regional.de.Saúde<-gsub("CEL. FABRICIANO","Coronel Fabriciano",ILTB_mapas$Unidade.Regional.de.Saúde)
ILTB_mapas$MUNIC_INSTITUICAO<-gsub("Gouvêa", "Gouveia",ILTB_mapas$MUNIC_INSTITUICAO)
ILTB_mapas$MUNIC_INSTITUICAO<-gsub("Queluzita", "Queluzito",ILTB_mapas$MUNIC_INSTITUICAO)
ILTB_mapas$MUNIC_INSTITUICAO<-gsub("São Thomé das Letras", "São Tomé das Letras",ILTB_mapas$MUNIC_INSTITUICAO)
ILTB_mapas$MUNIC_INSTITUICAO<-gsub("Brasópolis", "Brazópolis",ILTB_mapas$MUNIC_INSTITUICAO)
ILTB_mapas$MUNIC_INSTITUICAO<-gsub("Passa-Vinte", "Passa Vinte",ILTB_mapas$MUNIC_INSTITUICAO)

#Acrescentar a população:
MUNICIPIOS_mapas<-MUNICIPIOS_POP[,c("CÓDIGO","POP_Municipio","POP_URS")]

ILTB_mapas<- merge(ILTB_mapas, MUNICIPIOS_mapas, by = c("CÓDIGO")) 



## EXportação do banco para alimentação do painel ILTB nas páginas "Perfil da população" e "Quantitativos" 14.363 obs e 25 variables
#write.csv(ILTB_mapas,"C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/IL-TB/Painel/ILTB_mapa.csv") 
#write.csv(ILTB_mapas,"C:/Users/x3996631/OneDrive - Secretaria de Estado de Saude de Minas Gerais/IL-TB/Painel/ILTB_mapa.csv")



####################################################  INDICADOR ENCERRAMENTO DOS CASOS    ####################################################

#Selecionar apenas os tratamentos não encerrados com o tempo de tratamento expirado
ILTB_expirado<-subset(ILTB,TT_EXPIRADO =="Expirado")
ILTB_expirado<-subset(ILTB_expirado,Tratamento =="Não finalizado")
ILTB_expirado<-ILTB_expirado[,c("Unidade.Regional.de.Saúde","MUNIC_INSTITUICAO","ano","Número do caso","DT_NOTIF","DT_NASC",
                                "Medicamentos","Tempo_TT")]

ILTB_expirado$DT_NOTIF<-format(ILTB_expirado$DT_NOTIF,"%d-%m-%Y")
ILTB_expirado$DT_NASC<-format(ILTB_expirado$DT_NASC,"%d-%m-%Y")


#Ordenar o banco de acordo com a regional, município e ano da notificação
ILTB_expirado<-ILTB_expirado%>% arrange(Unidade.Regional.de.Saúde,MUNIC_INSTITUICAO,ano)

names(ILTB_expirado)[names(ILTB_expirado)=="Unidade.Regional.de.Saúde"]<-"Unidade Regional de Saúde"
names(ILTB_expirado)[names(ILTB_expirado)=="MUNIC_INSTITUICAO"]<-"Município"
names(ILTB_expirado)[names(ILTB_expirado)=="ano"]<-"Ano do início do tratamento" 
names(ILTB_expirado)[names(ILTB_expirado)=="Número do caso"]<-"Número da notificação"
names(ILTB_expirado)[names(ILTB_expirado)=="DT_NOTIF"]<-"Data da notificação"
names(ILTB_expirado)[names(ILTB_expirado)=="DT_NASC"]<-"Data de nascimento"
names(ILTB_expirado)[names(ILTB_expirado)=="Medicamentos"]<-"Tratamento prescrito"
names(ILTB_expirado)[names(ILTB_expirado)=="Tempo_em_TT"]<-"Tempo de tratamento (meses)"


##  EXPORTAÇÃO BANCO PARA A PLANILHA MOT_ILTB, ABA ENCERRAMENTOS    ## 710 obs e 8 variables
#write.xlsx(ILTB_expirado,"C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/IL-TB/Monitoramento/CASOS_SEM_ENCERRAMENTOS.xlsx") 
#write.xlsx(ILTB_expirado,"C:/Users/x3996631/OneDrive - Secretaria de Estado de Saude de Minas Gerais/IL-TB/Monitoramento/CASOS_SEM_ENCERRAMENTOS.xlsx") 

  

#################################################### INDICADOR DUPLICIDADES E REENTRADAS   ####################################################

  #Criação da variável ID, para exclusão das duplicidades nas notificações
ILTB$ID <- seq_len(nrow(ILTB))
  #Selecionar os registros onde Nome de registro = Data de nascimento 
ILTB_duplicadas <- ILTB %>%
  filter(duplicated(paste(Paciente,DT_NASC)) | duplicated(paste(Paciente,DT_NASC), fromLast = TRUE))

    # Excluir as duplicidades com pelo menos uma delas com dígito 2
names(ILTB_duplicadas)[names(ILTB_duplicadas)=="Dígito do caso"]<-"Digito"
names(ILTB_duplicadas)[names(ILTB_duplicadas)=="Número do caso"]<-"Numero"
ILTB_duplicadas_1<- ILTB_duplicadas %>%
  group_by(Paciente) %>%
  filter(any(Digito == 2)) %>%
  ungroup()
      # vetor para a exclusão
ILTB_duplicadas_2<- unique(ILTB_duplicadas_1$Numero) 
      #ExClusão das notificações com dígito 2
ILTB_duplicadas<-ILTB_duplicadas[!ILTB_duplicadas$Numero %in% ILTB_duplicadas_2,] 

#criação de código ID para exclusão do nome no banco final
ILTB_duplicadas <- ILTB_duplicadas%>%
  group_by(Paciente) %>%
  mutate(Identificacao_Paciente = cur_group_id()) %>%
  ungroup()

#Ordenar o banco de acordo com a Identificacao_Paciente e a data da notificação:
ILTB_duplicadas<-ILTB_duplicadas%>% arrange(Identificacao_Paciente,DT_NOTIF)

# Formatação datas:
ILTB_duplicadas$DT_NOTIF<-format(ILTB_duplicadas$DT_NOTIF,"%d-%m-%Y")
ILTB_duplicadas$DT_NASC<-format(ILTB_duplicadas$DT_NASC,"%d-%m-%Y")
ILTB_duplicadas$DT_ENCERRAMENTO<-format(ILTB_duplicadas$DT_ENCERRAMENTO,"%d-%m-%Y")

# Alteração de nomes das variáveis
names(ILTB_duplicadas)[names(ILTB_duplicadas)=="Identificacao_Paciente"]<-"Identificação do Paciente"
names(ILTB_duplicadas)[names(ILTB_duplicadas)=="Numero"]<-"Numero da notificação"
names(ILTB_duplicadas)[names(ILTB_duplicadas)=="DT_NOTIF"]<-"Data da notificação"
names(ILTB_duplicadas)[names(ILTB_duplicadas)=="DT_NASC"]<-"Data de nascimento do paciente"
names(ILTB_duplicadas)[names(ILTB_duplicadas)=="MUNIC_INSTITUICAO"]<-"Município"
names(ILTB_duplicadas)[names(ILTB_duplicadas)=="Unidade.Regional.de.Saúde"]<-"Unidade Regional de Saúde"
names(ILTB_duplicadas)[names(ILTB_duplicadas)=="DT_ENCERRAMENTO"]<-"Data do encerramento da notificação"
names(ILTB_duplicadas)[names(ILTB_duplicadas)=="ENCERRAMENTO"]<-"Motivo para encerramento"
names(ILTB_duplicadas)[names(ILTB_duplicadas)=="INDIC_TT_cat"]<-"Indicação de Tratamento"

#Seleção das variáveis
ILTB_duplicadas_final<-ILTB_duplicadas[,c("Unidade Regional de Saúde","Município","Identificação do Paciente","Data de nascimento do paciente",
                                          "Numero da notificação","Data da notificação","Tipo de entrada",
                                          "Indicação de Tratamento","Medicamentos","Data do encerramento da notificação",
                                          "Motivo para encerramento")]


##  EXPORTAÇÃO BANCO  MOT_ILTB, ABA DUPLICIDADES E REENTRADAS    ## 132 obs e 11 variables
#write.xlsx(ILTB_duplicadas_final,"C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/IL-TB/Monitoramento/ILTB_duplicidades_reentradas.xlsx") 
#write.xlsx(ILTB_duplicadas_final,"C:/Users/x3996631/OneDrive - Secretaria de Estado de Saude de Minas Gerais/IL-TB/Monitoramento/ILTB_duplicidades_reentradas.xlsx") 





##########################################   INDICADOR USO DE 3 HP   ##########################################################


## SELEÇÃO DAS VARIÁVEIS
ILTB_3HP<-ILTB[,c("Número do caso","DT_NASC","Medicamentos","Unidade.Regional.de.Saúde",
                  "MUNIC_INSTITUICAO","Gestante","IDADEcat","Tipo de entrada")]

## Selecionar apenas os tratamentos a partir de 2025  ##
ILTB_3HP<-subset(ILTB,DT_TT>="2025-01-1")

#Excluir gestantes, superior a 2 anos, uso de 3RH e reentradas após mudança de esquema ou após suspensão por condição clínica desfavorável ao tratamento;
ILTB_3HP<-subset(ILTB_3HP,Gestante!="Sim")
ILTB_3HP<-subset(ILTB_3HP,IDADEcat!="Inferior a 2 anos")
ILTB_3HP<-subset(ILTB_3HP,Medicamentos!="Rifampicina + Isoniazida - 3RH (dispersíveis pediátricos)")
ILTB_3HP<-subset(ILTB_3HP,`Tipo de entrada`!="Reentrada após mudança de esquema" &
                   `Tipo de entrada`!="Reentrada após suspensão por condição clínica desfavorável ao tratamento" )

##  CRIAÇÃO DE VARIÁVEIS  ##
#Total de tratamentos por regional
ILTB_3HP <- ILTB_3HP %>%
  group_by(Unidade.Regional.de.Saúde) %>%
  mutate(Total_TT_regional = n()) %>%
  ungroup()
#Total de tratamentos por município
ILTB_3HP <- ILTB_3HP %>%
  group_by(MUNIC_INSTITUICAO) %>%
  mutate(Total_TT_municip = n()) %>%
  ungroup()
#Total de tratamentos 3HP por regional
ILTB_3HP <- ILTB_3HP %>%
  group_by(Unidade.Regional.de.Saúde) %>%
  mutate(Total_HP3_regional = sum(Medicamentos == "Rifapentina + Isoniazida - 3HP", na.rm = TRUE)) %>%
  ungroup()
#Total de tratamentos 3HP por município
ILTB_3HP <- ILTB_3HP %>%
  group_by(MUNIC_INSTITUICAO) %>%
  mutate(Total_HP3_municip = sum(Medicamentos == "Rifapentina + Isoniazida - 3HP", na.rm = TRUE)) %>%
  ungroup()
#Proporção de tratamentos 3HP por regional
ILTB_3HP<- ILTB_3HP %>%
  mutate(Prop_3HP_Regional=floor((ILTB_3HP$Total_HP3_regional/ILTB_3HP$Total_TT_regional)*100))
#Proporção de tratamentos 3HP por município
ILTB_3HP<- ILTB_3HP %>%
  mutate(Prop_3HP_municip=floor((ILTB_3HP$Total_HP3_municip/ILTB_3HP$Total_TT_municip)*100))


## SELECIONAR E ORDENAR AS COLUNAS ##
ILTB_3HP<-ILTB_3HP[,c("Unidade.Regional.de.Saúde","MUNIC_INSTITUICAO",
                      "Número do caso","DT_NASC","Medicamentos","Total_TT_regional","Total_HP3_regional",
                      "Prop_3HP_Regional","Total_TT_municip","Total_HP3_municip","Prop_3HP_municip")]

#Ordenar banco pela regional
ILTB_3HP<-ILTB_3HP [order(ILTB_3HP$Unidade.Regional.de.Saúde),]  

##  MUDANÇA DO NOME DAS VARIÁVEIS ##
names(ILTB_3HP)[names(ILTB_3HP) == "Unidade.Regional.de.Saúde"] <- "Regional de Saúde"
names(ILTB_3HP)[names(ILTB_3HP) == "MUNIC_INSTITUICAO"] <- "Município"
names(ILTB_3HP)[names(ILTB_3HP) == "Número do caso"] <- "Número da notificação no IL-TB"
names(ILTB_3HP)[names(ILTB_3HP) == "DT_NASC"] <- "Data de nascimento do do Paciente"
names(ILTB_3HP)[names(ILTB_3HP) == "Medicamentos"] <- "Esquema terapêutico"
names(ILTB_3HP)[names(ILTB_3HP) == "Total_TT_regional"] <- "Total de tratamentos por Regional"
names(ILTB_3HP)[names(ILTB_3HP) == "Total_TT_municip"] <- "Total de tratamentos por município"
names(ILTB_3HP)[names(ILTB_3HP) == "Total_HP3_regional"] <- "Total de tratamentos 3HP por regional"
names(ILTB_3HP)[names(ILTB_3HP) == "Total_HP3_municip"] <- "Total de tratamentos 3HP por município"
names(ILTB_3HP)[names(ILTB_3HP) == "Prop_3HP_Regional"] <- "Proporção de tratamentos 3HP por Regional"
names(ILTB_3HP)[names(ILTB_3HP) == "Prop_3HP_municip"] <- "Proporção de tratamentos 3HP por município"


##  EXPORTAÇÃO para alimentar a planilha MOT- - 682 obs e 11 variables
#write.xlsx(ILTB_3HP,"C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/TBC/Monitoramento_oportuno_TB/Bancos_tratados/Indicador_3HP.xlsx")
#write.xlsx(ILTB_3HP,"C:/Users/x3996631/OneDrive - Secretaria de Estado de Saude de Minas Gerais/TBC/Monitoramento_oportuno_TB/Bancos_tratados/Indicador_3HP.xlsx")

## EXportação do banco para alimentação do painel ILTB
#write.csv(ILTB_3HP,"C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/IL-TB/Painel/ILTB_3HP.csv") 
#write.csv(ILTB_3HP,"C:/Users/x3996631/OneDrive - Secretaria de Estado de Saude de Minas Gerais/IL-TB/Painel/ILTB_3HP.csv")




##########################################   INDICADOR BUSCA ATIVA ÀS PVHA PARA O TRATAMENTO PREVENTIVO DE TUBERCULOSE   ###################################

## SELEÇÃO DAS VARIÁVEIS ## 
SIMC<-SIMC_bruto[,c("Instituicao.solicitante","Município.instituição.solicitante",
                    "Paciente","Dt.Nascimento","Ação.realizada","Dt..atualização","Valor.do.CD4.na.ocasião","Data.do.exame.de.contagem.de.CD4")]

#Preparo da variável "Município.instituição.solicitante" 
SIMC$Município.instituição.solicitante<-gsub("Ibirite", "ibirité",SIMC$Município.instituição.solicitante)
SIMC$Município.instituição.solicitante<-toupper(SIMC$Município.instituição.solicitante)

## SELEÇÃO DOS MUNICÍPIOS DE MINAS GERAIS ##
MUNICIPIOS_HIV<-MUNICIPIOS[,c("NM_MUNICIP","Unidade.Regional.de.Saúde")]
names(MUNICIPIOS_HIV)[names(MUNICIPIOS_HIV)=="NM_MUNICIP"]<- "Município.instituição.solicitante"
MUNICIPIOS_HIV$Município.instituição.solicitante<-toupper(MUNICIPIOS_HIV$Município.instituição.solicitante)


## VARIÁVEL "Unidade.Regional.de.Saúde"
SIMC<- merge(SIMC, MUNICIPIOS_HIV, by = c("Município.instituição.solicitante"))

## Preparo das variáveis ##
SIMC$Paciente<-toupper(SIMC$Paciente)
SIMC$Dt.Nascimento <- as.Date(as.numeric(SIMC$Dt.Nascimento), origin = "1899-12-30")
SIMC$Dt..atualização<- format(as.Date(as.numeric(SIMC$Dt..atualização), origin = "1899-12-30"), "%d-%m-%Y")
SIMC$Data.do.exame.de.contagem.de.CD4 <- as.Date(as.numeric(SIMC$Data.do.exame.de.contagem.de.CD4), origin = "1899-12-30")


## EXCLUSÕES ##
#Excluir CD4 superior a 350
SIMC<-subset(SIMC,Valor.do.CD4.na.ocasião<350)
#Excluir os serviços de saúde que não fazem parte de Minas Gerais, municípios homônimos
SIMC<-subset(SIMC,Instituicao.solicitante!="SERVIÇO DE ASSISTÊNCIA ESPECIALIZADA DE ÁGUA BOA")
SIMC<-subset(SIMC,Instituicao.solicitante!="CONSÓRCIO INTERMUNICIPAL DE SAÚDE DA COSTA OESTE DO PARANÁ")
SIMC<-subset(SIMC,Instituicao.solicitante!="SAE - SERVIÇO DE ASSISTÊNCIA ESPECIALIZADA - ITAPEVA")
SIMC<-subset(SIMC,Instituicao.solicitante!="Posto de Saúde de Mesquita")
SIMC<-subset(SIMC,Instituicao.solicitante!="FUNDO MUNICIPAL DE SAUDE DE BOA ESPERANÇA")
SIMC<-subset(SIMC,Instituicao.solicitante!="Secretaria Municipal de Saúde de Cantagalo - RJ")


#Exclusão dos pacientes notificados para tuberculose:
#Considerar os tratamentos de TB realizados no mesmo período do banco SIMC: ano anterior e no ano atual: 

SINAN_TB<-subset(SINAN_TB_bruto,NU_ANO=="2025" | NU_ANO=="2026")
SINAN_TB<-SINAN_TB[,c("NM_PACIENT","DT_NASC")]
names(SINAN_TB)[names(SINAN_TB)=="NM_PACIENT"]<-"Paciente1"
names(SINAN_TB)[names(SINAN_TB)=="DT_NASC"]<-"Dt.Nascimento"
SINAN_TB$Paciente1<-as.character(SINAN_TB$Paciente1)
SINAN_TB<-unique(SINAN_TB)

#Preparo da variável "Paciente" pra a linkagem SINAN X SIMC:
#No SIMC
SIMC$Paciente1 <- iconv(SIMC$Paciente, from = "UTF-8", to = "ASCII//TRANSLIT") #exclui acentos e Ç
SIMC$Paciente1 <- gsub("[[:punct:]]", "", SIMC$Paciente1) #exclui pontuação
SIMC$Paciente1 <- gsub(" ","", SIMC$Paciente1) #exclui espaço

##No SINAN
SINAN_TB$Paciente1 <- iconv(SINAN_TB$Paciente1, from = "UTF-8", to = "ASCII//TRANSLIT") 
SINAN_TB$Paciente1 <- gsub("[[:punct:]]", "", SINAN_TB$Paciente1) 
SINAN_TB$Paciente1 <- gsub(" ","", SINAN_TB$Paciente1) 

SIMC<- anti_join(SIMC, SINAN_TB, by = c("Paciente1","Dt.Nascimento"))


#Exclusão dos pacientes notificados para ILTB:
  #Considerar os tratamentos de ILTB realizados no mesmo período do banco SIMC: ano anterior e no ano atual: 
ILTB_SIMC<-subset(ILTB,DT_NOTIF>="2025-01-01")
ILTB_SIMC<-ILTB_SIMC[,c("Paciente","DT_NASC")]
names(ILTB_SIMC)[names(ILTB_SIMC)=="DT_NASC"]<-"Dt.Nascimento"
names(ILTB_SIMC)[names(ILTB_SIMC)=="Paciente"]<-"Paciente1"
ILTB_SIMC<-unique(ILTB_SIMC)

#Preparo da variável "Paciente" pra a linkagem SINAN X SIMC:
ILTB_SIMC$Paciente1 <- iconv(ILTB_SIMC$Paciente, from = "UTF-8", to = "ASCII//TRANSLIT") #exclui acentos e Ç
ILTB_SIMC$Paciente1 <- gsub("[[:punct:]]", "", ILTB_SIMC$Paciente1) #exclui pontuação
ILTB_SIMC$Paciente1 <- gsub(" ","", ILTB_SIMC$Paciente1) #exclui espaço

SIMC<- anti_join(SIMC, ILTB_SIMC, by = c("Paciente1","Dt.Nascimento"))


## Ação realizada  ##
# Preencher "como"REALIZAR BUSCA ATIVA" nos NAs: 
names(SIMC)[names(SIMC)=="Ação.realizada"]<-"Acao.realizada"
SIMC <- SIMC %>%
  mutate(Acao.realizada = if_else(is.na(Acao.realizada) | Acao.realizada == "", "Realizar busca ativa", Acao.realizada))


## CONSTRUÇÃO DE VARIÁVEIS  ##

# Variável tipo de atendimento: Hospitalar ou ambulatorial:Utilizar esta variável para criar a ação realizada Indicar LF_LAM
SIMC<- SIMC%>%
  mutate(Atendimento = case_when(
    Instituicao.solicitante %in% c("Ambulatório de Doenças Infecciosas e Parasitárias Ribeirão das Neves",
                                   "Ambulatório de Infectologia de São Sebastião do Paraíso",
                                   "Ambulatório de MI Herbert de Souza - [cp]",
                                   "Ambulatório Escola da FAENPA / Passos",
                                   "Ambulatório Médico Especializado de Lavras",
                                   "ASSOCIACAO MARIO PENNA",
                                   "BONFIM - UBS DIMINIZ DINIZ DA SILVA",
                                   "Centro Ambulatorial de Especialidades Tancredo Neves -CAETAN",
                                   "Centro de Apoio Especializado em DST/HIV/AIDS de Araguari",
                                   "Centro de Assistência e Prevenção DST/AIDS de Itajubá",
                                   "Centro de Consultas Especializadas Iria Diniz",
                                   "Centro de Promoção Cristiano Azevedo",
                                   "CENTRO DE PROMOÇÃO DA SAÚDE- CTA/SAE- CONSELHEIRO LAFAIETE",
                                   "Centro de Referência e Atenção Especial à Saúde - CRASE",
                                   "Centro De Referência Em Doenças Infecciosas- CEREDI",
                                   "Centro de Referência em Doenças Infecciosas - Montes Claros",
                                   "Centro de Referência Vital Brazil - Vespasiano",
                                   "Centro de Saúde Aldo Olivotti de Extrema",
                                   "Centro de Saúde Padre Hildebrando de Freitas",
                                   "Centro de Saúde Santa Rita de Sapucaí",
                                   "CENTRO DE TESTAGEM E ACOLHIMENTO DE MANTENA",
                                   "Centro de Testagem e Aconselhamento de Araxá - MG",
                                   "Centro de Testagem e Aconselhamento de Pouso Alegre-SAE",
                                   "CENTRO DE TESTAGEM E ACONSELHAMENTO EM DST/HIV/AIDS ÁGUAS FORMOSAS",
                                   "Centro de Tratamento e Referência Orestes Diniz",
                                   "Centro de Vigilância em Saúde de Formiga",
                                   "Centro Especializado de Assistência a Saúde Comunitária - UNIFENAS",
                                   "Centro Municipal de Saúde - Paraisópolis",
                                   "CEREDI DE JANAÚBA",
                                   "CTA-SAE Sagrada Família",
                                   "CTA DE PATOS DE MINAS - CENTRO VIVA VIDA DONA FRANCISCA ESCOLÁSTICA",
                                   "CTA/SAE DE ITAOBIM",
                                   "CTR - Vale do Aço - Policlínica Municipal de Ipatinga-MG",
                                   "Diretoria Regional de Saúde de Ubá",
                                   "Fundação Municipal de Saúde de São Lourenço - PM de DST/Aids",
                                   "Laboratório Municipal de Santa Luzia",
                                   "Núcleo de Apoio Psicossocial de Três Pontas - NAPS",
                                   "Núcleo Especializado em Programas de Saude - NEPS",
                                   "Policlínica de Referência de Barbacena",
                                   "Policlínica de Varginha",
                                   "Policlínica Milton Campos",
                                   "POLICLÍNICA MUNICIPAL AFONSINA NUNES ARAÚJO - ARAÇUAI",
                                   "POLICLÍNICA MUNICIPAL DE CARANGOLA",
                                   "Policlínica Municipal de Itabira",
                                   "Policlínica Municipal de Ituiutaba",
                                   "Policlínica Municipal de Saúde de Divinópolis",
                                   "Policlínica Municipal de Teófilo Otoni",
                                   "POLICLINICA MUNICIPAL DR. JORGE HANNAS",
                                   "Policlínica Regional de Diamantina",
                                   "Programa Municipal de DST/Aids de Poços de Caldas",
                                   "SAE - Itaúna",
                                   "SAE - Ouro Preto",
                                   "SAE - SABARÁ",
                                   "SAE  -  IBIRITÉ",
                                   "SAE/CTA/UDM Centro de Saúde Paulo Vilela Loureiro",
                                   "Secretaria de Saúde de Laranjal",
                                   "Secretaria Municipal de Saúde de Alfenas",
                                   "Secretaria Municipal de Saúde de Cataguases",
                                   "Secretaria Municipal de Saúde de Itabirito",
                                   "Secretaria Municipal de Saúde de Muriaé",
                                   "Secretaria Municipal de Saúde de Santos Dumont",
                                   "Secretaria Municipal de Saúde de São João del Rei",
                                   "Secretaria Municipal de Saúde de Timóteo - MG",
                                   "Secretaria Municipal de Saúde de Três Corações",
                                   "SECRETARIA MUNICIPAL DE SAUDE DE VISCONDE DO RIO BRANCO",
                                   "SEPADI- PREFEITURA MUNICIPAL DE BETIM",
                                   "SERVIÇO ATENDIMENTO MÉDICO ESPECIALIZADO - SAME DE UNAÍ-MG",
                                   "Serviço de Assistência Especializada - Ambulatório de DST/Aids da Prefeitura Municipal de Frutal",
                                   "SERVIÇO DE ASSISTÊNCIA ESPECIALIZADA - NOVA LIMA",
                                   "Serviço de Assistência Especializada de Juiz de Fora",
                                   "Serviço de Assistência Especializada de Uberaba - SAE - CTA - Prefeitura Municipal [cp]",
                                   "Serviço de Atenção Especializada Dr. Marcio Tadeu Diniz de Souza",
                                   "SERVIÇO DE ATENÇÃO ESPECIALIZADA SAE - CARATINGA",
                                   "UDM - Além Paraíba",
                                   "Unidade de Atendimento Especializado de Viçosa",
                                   "Unidade de Referência Secundária Centro Sul/PBH",
                                   "Universidade Federal de Uberlândia - Departamento de Doenças Infecto-Contagiosas [cp]",
                                   "Universidade Federal do Triângulo Mineiro - Lab. de Uberaba #","Policlínica de Andradas",
                                   "Secretaria Municipal de Araxá - CTA","UBSS Serviço de Assistência Especializada – Nova Lima",
                                   "CENTRO DE SAÚDE CRISTINO JOSÉ DA SILVA") ~ "Ambulatorial",
    Instituicao.solicitante %in% c("HOSPITAL DE PRONTO SOCORRO DOUTOR MOZART GERALDO TEIXEIRA",
                                   "HOSPITAL E MATERNIDADE THEREZINHA DE JESUS",
                                   "Hospital Eduardo de Menezes - FHEMIG",
                                   "HOSPITAL MUNICIPAL DE CONTAGEM",
                                   "HOSPITAL MUNICIPAL VALDEMAR DE ASSIS BARCELOS",
                                   "HOSPITAL ODILON BEHRENS",
                                   "HOSPITAL REGIONAL JOÃO PENIDO",
                                   "HOSPITAL UNIVERSITÁRIO CIÊNCIAS MÉDICAS",
                                   "Hospital Universitário de Juiz de Fora",
                                   "Laboratório de Patologia Clínica Hospital João XXIII",
                                   "Serviço Social Autônomo do Hospital Metropolitano Dr. Célio de Castro",	
                                    "HOSPITAL MUNICIPAL PEDRO DOS REIS FERNANDES NETO - SANTA LUZIA/MA",
                                   "HOSPITAL PADRE JULIO MARIA") ~ "Hospitalar",
    TRUE ~ "Ignorado"))
#Verificar a presença da categoria "Ignorado", se sim atualizar comando acima  
table(SIMC$Atendimento)


#Criar a categoria Realizar exame LF_LAM na variável "Ação_realizada" :
SIMC <- SIMC %>%
  mutate(Acao.realizada = case_when(
    Atendimento == "Ambulatorial" & Valor.do.CD4.na.ocasião <= 100 & Acao.realizada== "Realizar busca ativa"~ "Oferecer teste LF-LAM" ,
    Atendimento == "Hospitalar"   & Valor.do.CD4.na.ocasião <= 200 & Acao.realizada== "Realizar busca ativa" ~ "Oferecer teste LF-LAM",
    Atendimento == "Ambulatorial" & Valor.do.CD4.na.ocasião > 100 & Acao.realizada== "Realizar busca ativa"~ "Oferecer tratamento preventivo TB" ,
    Atendimento == "Hospitalar"   & Valor.do.CD4.na.ocasião > 200 & Acao.realizada== "Realizar busca ativa" ~ "Oferecer tratamento preventivo TB",
    Acao.realizada %in% c( "Indicado tratamento para ILTB","Indicado tratamento para TB","Iniciado tratamento da ILTB e notificado no IL-TB") ~ "Encaminhado para Tratamento TB/ILTB",
    Acao.realizada %in% c ("Óbito","Óbito identificado automaticamente pelo sistema") ~"Òbito",
    Acao.realizada %in% c ("Pessoa já tratou ILTB no passado","Pessoa já tratou TB no passado") ~"Pessoa tratou TB/ILTB no passado",
    TRUE ~ Acao.realizada
  ))
                        
  
#Total de pacientes imunossuprimidos por regional#:
SIMC <- SIMC %>%
  group_by(Unidade.Regional.de.Saúde) %>%
  mutate(Tot_pac_Regional = n()) %>%
  ungroup()
#Total de pacientes imunossuprimidos por municípios#
SIMC <- SIMC %>%
  group_by(Município.instituição.solicitante) %>%
  mutate(Tot_pac_Municipio = n()) %>%
  ungroup()
#Total de pacientes imunossuprimidos por regional sem busca ativa#
SIMC <- SIMC %>%
  group_by(Unidade.Regional.de.Saúde) %>%
  mutate(Tot_pac_Regional_busca_ativa = sum(Acao.realizada == "REALIZAR BUSCA ATIVA", na.rm = TRUE)) %>%
  ungroup()
#Total de pacientes imunossuprimidos por município sem busca ativa#
SIMC <- SIMC %>%
  group_by(Município.instituição.solicitante) %>%
  mutate(Tot_pac_Municip_busca_ativa = sum(Acao.realizada == "REALIZAR BUSCA ATIVA", na.rm = TRUE)) %>%
  ungroup()
#Proporção de pacientes sem busca ativa por regional:
SIMC <- SIMC %>%
  mutate(Proporção_regional = floor((Tot_pac_Regional_busca_ativa / Tot_pac_Regional) * 100))
#Proporção de pacientes sem busca ativa por Município:
SIMC<- SIMC %>%
  mutate(Proporção_municip = floor((Tot_pac_Municip_busca_ativa/Tot_pac_Municipio)*100))


## Selecionar e ordenar as colunas ##
SIMC<-SIMC[,c("Unidade.Regional.de.Saúde","Município.instituição.solicitante","Instituicao.solicitante",
              "Paciente","Dt.Nascimento","Valor.do.CD4.na.ocasião","Data.do.exame.de.contagem.de.CD4","Acao.realizada","Dt..atualização","Tot_pac_Regional","Tot_pac_Regional_busca_ativa",
              "Proporção_regional","Tot_pac_Municipio","Tot_pac_Municip_busca_ativa","Proporção_municip")]

#Ordenar banco pela regional
SIMC <- SIMC [order(SIMC$Unidade.Regional.de.Saúde),]  

##  MUDANÇA DO NOME DAS VARIÁVEIS ##
names(SIMC)[names(SIMC) == "Unidade.Regional.de.Saúde"] <- "Regional de Saúde"
names(SIMC)[names(SIMC) == "Município.instituição.solicitante"] <- "Município"
names(SIMC)[names(SIMC) == "Instituicao.solicitante"] <- "Unidade de Saúde"
names(SIMC)[names(SIMC) == "Paciente"] <- "Nome do Paciente"
names(SIMC)[names(SIMC) == "Dt.Nascimento"] <- "Data de nascimento do Paciente"
names(SIMC)[names(SIMC) == "Valor.do.CD4.na.ocasião"] <- "Última contagem de CD4"
names(SIMC)[names(SIMC) == "Data.do.exame.de.contagem.de.CD4"] <- "Data do exame CD4"
names(SIMC)[names(SIMC) == "Acao.realizada"] <- "Ação realizada"
names(SIMC)[names(SIMC) == "Dt..atualização"] <- "Data de atualização no sistema"
names(SIMC)[names(SIMC) == "Tot_pac_Regional"] <- "Total de pacientes imunossuprimidos por Regional"
names(SIMC)[names(SIMC) == "Tot_pac_Regional_busca_ativa"] <- "Total de pacientes imunossuprimidos por Regional para busca ativa"
names(SIMC)[names(SIMC) == "Proporção_regional"] <- "Proporção de pacientes imunossuprimidos por Regional para busca ativa"
names(SIMC)[names(SIMC) == "Tot_pac_Municipio"] <- "Total de pacientes imunossuprimidos por Municípios"
names(SIMC)[names(SIMC) == "Tot_pac_Municip_busca_ativa"] <- "Total de pacientes imunossuprimidos por Municípios para busca ativa"
names(SIMC)[names(SIMC) == "Proporção_municip"] <- "Proporção de pacientes imunossuprimidos por Município para busca ativa"


#Seleção para a planilha MOT-TB:
SIMC_MOT<-SIMC[,c("Regional de Saúde","Município","Unidade de Saúde","Nome do Paciente","Data de nascimento do Paciente","Última contagem de CD4",
              "Data do exame CD4","Ação realizada","Data de atualização no sistema")]

SIMC_Painel<-SIMC[,c("Regional de Saúde","Município","Unidade de Saúde","Nome do Paciente","Data de nascimento do Paciente","Ação realizada",
                     "Data de atualização no sistema","Total de pacientes imunossuprimidos por Regional","Total de pacientes imunossuprimidos por Regional para busca ativa",
                     "Proporção de pacientes imunossuprimidos por Regional para busca ativa","Total de pacientes imunossuprimidos por Municípios",
                     "Total de pacientes imunossuprimidos por Municípios para busca ativa","Proporção de pacientes imunossuprimidos por Município para busca ativa")]


#Exportação para planilha MOT_TB
#write.xlsx(SIMC_MOT,"C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/TBC/Monitoramento_oportuno_TB/Bancos_tratados/INDICARDOR_BUSCA_ATIVA.xlsx") 
#write.xlsx(SIMC_MOT,"C:/Users/x3996631/OneDrive - Secretaria de Estado de Saude de Minas Gerais/TBC/Monitoramento_oportuno_TB/Bancos_tratados/INDICARDOR_BUSCA_ATIVA.xlsx")

## EXportação do banco para alimentação do painel ILTB
#write.csv(SIMC_Painel,"C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/IL-TB/Painel/ILTB_HIV.csv") 
#write.csv(SIMC_Painel,"C:/Users/x3996631/OneDrive - Secretaria de Estado de Saude de Minas Gerais/IL-TB/Painel/ILTB_HIV.csv")




##########################################   INDICADOR DESCENTRALIZAÇÃO DO SISTEMA IL-TB ##########################################################


## SELEÇÃO DAS VARIÁVEIS
USUARIOS<-USUARIOS_BRUTO[,c("Transações nos últimos 12 meses","Município (Cod. IBGE)")]
MUNICIPIOS_desc<-MUNICIPIOS_POP[,c("Unidade.Regional.de.Saúde","NM_MUNICIP","CÓDIGO","POP_Municipio")]

##  Seleção dos municípios com mais de 10.000 habitantes:
MUNICIPIOS_desc<-subset(MUNICIPIOS_desc,POP_Municipio>=10000)

## VARIÁVEL MUNICÍPIO, REGIONAL E POPULAÇÃO  ##
names(USUARIOS)[names(USUARIOS) == "Município (Cod. IBGE)"] <- "CÓDIGO"
USUARIOS<- merge(USUARIOS, MUNICIPIOS_desc, by = c("CÓDIGO"), all=TRUE)

####  CRIAÇÃO DE VARIÁVEIS  ##
#Total  de transações por município
USUARIOS <- USUARIOS %>%
  group_by(NM_MUNICIP) %>%
  mutate(Total_acoes_município = sum(`Transações nos últimos 12 meses`, na.rm = TRUE)) %>%
  ungroup()
USUARIOS<-USUARIOS[,c("Unidade.Regional.de.Saúde","NM_MUNICIP","Total_acoes_município")]
USUARIOS<-unique(USUARIOS)

#Total de transações por município >10.000 por regional
USUARIOS <- USUARIOS %>%
  group_by(Unidade.Regional.de.Saúde) %>%
  mutate(Total_municip = n()) %>%
  ungroup()

# Total de municípios sem ação no IL-TB:
USUARIOS <- USUARIOS %>%
  group_by(Unidade.Regional.de.Saúde) %>%
  mutate(Total_municip_sem_acao = sum (Total_acoes_município == 0, na.rm = TRUE)) %>%
  ungroup()

##Proporção de municípios sem ação no IL-TB
USUARIOS$Prop_municip_sem_acao<-round((USUARIOS$Total_municip_sem_acao/USUARIOS$Total_municip)*100,0)

## SELECIONAR E ORDENAR AS COLUNAS ##
USUARIOS<-USUARIOS[,c("Unidade.Regional.de.Saúde","NM_MUNICIP","Total_municip","Total_acoes_município",
                      "Total_municip_sem_acao","Prop_municip_sem_acao")]

##  ORDENAR O BANCO
USUARIOS<- USUARIOS %>%
  arrange(Unidade.Regional.de.Saúde)

##  MUDANÇA DO NOME DAS VARIÁVEIS ##
names(USUARIOS)[names(USUARIOS) == "Unidade.Regional.de.Saúde"] <- "Regional de Saúde"
names(USUARIOS)[names(USUARIOS) == "NM_MUNICIP"] <- "Município com mais de 10.000 habitantes"
names(USUARIOS)[names(USUARIOS) == "Total_municip"] <- "Total de Municípios com mais de 10.000 habitantes por regional"
names(USUARIOS)[names(USUARIOS) == "Total_acoes_município"] <- "Total de ações no Sistema IL-TB do município"
names(USUARIOS)[names(USUARIOS) == "Total_municip_sem_acao"] <- "Total de municípios que não utilizaram o sistema IL-TB por regional"
names(USUARIOS)[names(USUARIOS) == "Prop_municip_sem_acao"] <- "Proporção de municípios que não utilizaram o sistema IL-TB por regional"



##  EXPORTAÇÃO
## EXportação do banco para alimentação da planilha MOT_ILTB
#write.xlsx(USUARIOS,"C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/TBC/Monitoramento_oportuno_TB/Bancos_tratados/Indicador_DESCENTRALIZAÇÃO.xlsx")
#write.xlsx(USUARIOS,"C:/Users/x3996631/OneDrive - Secretaria de Estado de Saude de Minas Gerais/TBC/Monitoramento_oportuno_TB/Bancos_tratados/Indicador_DESCENTRALIZAÇÃO.xlsx")

## EXportação do banco para alimentação do painel ILTB
#write.csv(USUARIOS,"C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/IL-TB/Painel/ILTB_descentralizacao.csv") 
#write.csv(USUARIOS,"C:/Users/x3996631/OneDrive - Secretaria de Estado de Saude de Minas Gerais/IL-TB/Painel/ILTB_descentralizacao.csv")




###################       BANCO MOT-ILTB ABA TRATAMENTOS CONTRAINDICADOS    ########################
# SERÃO CONSIDERADOS TRATAMENTOS CONTRAINDICADOS:
#Gestantes em uso de 3HP
#Crianças abaixo de 2 anos em uso de 3HP
#Indivíduos acima de 50 anos em uso de 6H ou 9H
# Indivíduos HIV em uso de rifampicina (NOTA INFORMATIVA Nº 33/2024-CGAFME/DAF/SECTICS/MS) FALTA COLOCAR NO SCRIPT, mas preciso saber o tipo de TARV

ILTB_tt_contraindicados<-subset(ILTB,Tratamento=="Não finalizado")

#Seleção das variáveis
ILTB_tt_contraindicados<-ILTB_tt_contraindicados[,c("Unidade.Regional.de.Saúde","MUNIC_INSTITUICAO","ano","Número do caso","DT_NOTIF","DT_NASC",
                                                    "Medicamentos","IDADEcat","Gestante")]


#Criação da variável "Tratamento contraindicado":
ILTB_tt_contraindicados <- ILTB_tt_contraindicados %>%
  mutate(
    TT_contraindicado = case_when(
      Gestante %in% c("Sim", "Ignorado") & Medicamentos== "Rifapentina + Isoniazida - 3HP" ~ "Gestante ou possibilidade gestação em uso de 3HP",
      IDADEcat == "Inferior a 2 anos" &  Medicamentos== "Rifapentina + Isoniazida - 3HP"~ "Criança abaixo 2 anos em uso de 3HP",
      IDADEcat %in% c("50 a 59 anos","Superior a 60 anos") &  Medicamentos %in% c("Isoniazida","Isoniazida - 6H","Isoniazida - 9H") ~ "Indivíduo acima de 50 anos em uso de isoniazida",
      TRUE ~ NA_character_ # Para valores não especificados, retorna NA
    )
  )

#Exclusão dos tratamentos não contraindicados
ILTB_tt_contraindicados<-subset(ILTB_tt_contraindicados,!is.na(TT_contraindicado))

#Mudança na formatação data
ILTB_tt_contraindicados$DT_NOTIF<-format(ILTB_tt_contraindicados$DT_NOTIF,"%d-%m-%Y")
ILTB_tt_contraindicados$DT_NASC<-format(ILTB_tt_contraindicados$DT_NASC,"%d-%m-%Y")

#Ordenar o banco de acordo com a regional, município e ano da notificação
ILTB_tt_contraindicados$ano<-as.numeric(ILTB_tt_contraindicados$ano)
ILTB_tt_contraindicados<-ILTB_tt_contraindicados%>% arrange(Unidade.Regional.de.Saúde,MUNIC_INSTITUICAO,ano)

names(ILTB_tt_contraindicados)[names(ILTB_tt_contraindicados)=="Unidade.Regional.de.Saúde"]<-"Unidade Regional de Saúde"
names(ILTB_tt_contraindicados)[names(ILTB_tt_contraindicados)=="MUNIC_INSTITUICAO"]<-"Município"
names(ILTB_tt_contraindicados)[names(ILTB_tt_contraindicados)=="ano"]<-"Ano do início do tratamento" 
names(ILTB_tt_contraindicados)[names(ILTB_tt_contraindicados)=="Número do caso"]<-"Número da notificação"
names(ILTB_tt_contraindicados)[names(ILTB_tt_contraindicados)=="DT_NOTIF"]<-"Data da notificação"
names(ILTB_tt_contraindicados)[names(ILTB_tt_contraindicados)=="DT_NASC"]<-"Data de nascimento"
names(ILTB_tt_contraindicados)[names(ILTB_tt_contraindicados)=="Medicamentos"]<-"Tratamento prescrito"
names(ILTB_tt_contraindicados)[names(ILTB_tt_contraindicados)=="IDADEcat"]<-"Idade"
names(ILTB_tt_contraindicados)[names(ILTB_tt_contraindicados)=="TT_contraindicado"]<-"Contraindicação"


## EXportação do banco para encaminhar para a SAF
#write.xlsx(ILTB_tt_contraindicados,"C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Medicamentos/Tratamentos contraindicados/Tratamentos_ILTB_contraindicados.xlsx")
#write.xlsx(ILTB_tt_contraindicados,"C:/Users/x3996631/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Medicamentos/Tratamentos contraindicados/Tratamentos_ILTB_contraindicados.xlsx")






############################################         PAREI AQUI      ###################################
### PAREAMENTO PROBABILÍSTICO: SIMC_MINAS X ILTB
##Melhorar o pareamento das datas, passar a utilizar o nome da mãe além nome paciente e data nasc
#Exportação dos bancos para tratamento no Python
#write.xlsx(SIMC_MG,"C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/TBC/Site SIMC/BAncos_tratados/SIMC_MG.xlsx")
#write.xlsx(ILTB,"C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/TBC/Site SIMC/BAncos_tratados/ILTB_tratado.xlsx")

###########################################################################################################


#################################################### Boletim 2024  #########################################

#A partir de 2018
ILTB$ano<-as.numeric(ILTB$ano)
ILTB_2018_2024<-subset(ILTB,ano>=2018 & ano<2025)

# TABELAS DESCRITIVAS

table(ILTB_2018_2024$Sexo)
sexo_porcentagem<- prop.table(table(ILTB_2018_2024$Sexo))*100
sexo_porcentagem
table(ILTB_2018_2024$Sexo,ILTB_2018_2024$ano)
sexo_porcentagem_ano <- prop.table(table(ILTB_2018_2024$Sexo, ILTB_2018_2024$ano), margin = 2) * 100
sexo_porcentagem_ano

#Recategorizar a idade:
ILTB_2018_2024$IDADEcat2<-rep(NA,length(ILTB_2018_2024$IDADE))
ILTB_2018_2024$IDADEcat2[ILTB_2018_2024$IDADE<=12]<-"0 - 12 anos" # criança
ILTB_2018_2024$IDADEcat2[12<ILTB_2018_2024$IDADE & ILTB_2018_2024$IDADE<=25]<-"13 - 25 anos" #adolescente/jovem
ILTB_2018_2024$IDADEcat2[25<ILTB_2018_2024$IDADE & ILTB_2018_2024$IDADE<=35]<-"26 - 35 anos" # adulto jovem
ILTB_2018_2024$IDADEcat2[35<ILTB_2018_2024$IDADE & ILTB_2018_2024$IDADE<=60]<-"36 - 60 anos" # adulto
ILTB_2018_2024$IDADEcat2[ILTB_2018_2024$IDADE>60]<-"60 anos ou mais" # Idoso
ILTB_2018_2024$IDADEcat2=as.factor(ILTB_2018_2024$IDADEcat2)
table(ILTB_2018_2024$IDADEcat2)
IDADEcat2_porcentagem<- prop.table(table(ILTB_2018_2024$IDADEcat2))*100
IDADEcat2_porcentagem
table(ILTB_2018_2024$IDADEcat2,ILTB_2018_2024$ano)
IDADEcat2_porcentagem_ano<- prop.table(table(ILTB_2018_2024$IDADEcat2, ILTB_2018_2024$ano), margin = 2) * 100
IDADEcat2_porcentagem_ano
summary(ILTB_2018_2024$Idade)
tapply(ILTB_2018_2024$Idade, ILTB_2018_2024$ano, summary)


#Recategorizar raça/cor
table(ILTB_2018_2024$`Raça/cor`)
ILTB_2018_2024 <- ILTB_2018_2024 %>% 
  mutate(raca_cor_cat = case_when(
    `Raça/cor` == "Parda" ~ "Parda",
    `Raça/cor` == "Preta" ~ "Preta",
    `Raça/cor` == "Branca" ~ "Branca",
    `Raça/cor` == "IGNORADO" ~ "IGNORADO",
    `Raça/cor` %in% c("Amarela", "Indígena") ~ "Amarela_Indígena",
    TRUE ~ "INDETERMINADA"
  ))
table(ILTB_2018_2024$raca_cor_cat)
#NÃO CONSIDERAR CATEGORIAS IGNORADA
raça_cor<-subset(ILTB_2018_2024,raca_cor_cat!="IGNORADO") 
table(raça_cor$raca_cor_cat)
raca_cor_porcentagem<- prop.table(table(raça_cor$raca_cor_cat))*100
raca_cor_porcentagem
table(raça_cor$raca_cor_cat,raça_cor$ano)
raca_cor_porcentagem_ano<- prop.table(table(raça_cor$raca_cor_cat,raça_cor$ano), margin = 2) * 100
raca_cor_porcentagem_ano


#INDICAÇÃO DO TRATAMENTO:recategorizar
ILTB_2018_2024 <- ILTB_2018_2024 %>% 
  mutate(
    Indicacao_TT = case_when(
      `Indicação de Tratamento` %in% c("Alterações radiológicas", "Profissionais de saúde/ILP",
                                       "Fumo/Silicose/baixo peso", "Outras condições", "Recém-nascidos") ~ "Outras condições",
      `Indicação de Tratamento` %in% c("Cond.crônicas/Neoplasia", "Uso de Imunossupressores") ~ "Imunossupressão",
      `Indicação de Tratamento` == "Contatos TB pulmonar" ~ "Contatos TB pulmonar",
      `Indicação de Tratamento` == "Pessoas vivendo com HIV" ~ "Pessoas vivendo com HIV",
      TRUE ~ "Indeterminado"
    )
  )

table(ILTB_2018_2024$Indicacao_TT)
TT_porcentagem<- prop.table(table(ILTB_2018_2024$Indicacao_TT))*100
TT_porcentagem
table(ILTB_2018_2024$Indicacao_TT,ILTB_2018_2024$ano)
TT_porcentagem_ano <- prop.table(table(ILTB_2018_2024$Indicacao_TT, ILTB_2018_2024$ano), margin = 2) * 100
TT_porcentagem_ano
 

## EXCLUIR OS CASOS DE TRATAMENTOS NÃO ENCERRADOS
ILTB_nao_encerrados<-subset(ILTB_2018_2024,TT_completo!="Tramento não encerrado")
table(ILTB_nao_encerrados$TT_completo)
Completo_porcentagem<- prop.table(table(ILTB_nao_encerrados$TT_completo))*100
Completo_porcentagem
table(ILTB_nao_encerrados$TT_completo,ILTB_nao_encerrados$ano)
Completo_porcentagem_ano <- prop.table(table(ILTB_nao_encerrados$TT_completo, ILTB_nao_encerrados$ano), margin = 2) * 100
Completo_porcentagem_ano

table(ILTB_2018_2024$ENCERRAMENTO)
#ENCERRAMENTO:

#CATEGORIZAR
ILTB_nao_encerrados<- ILTB_nao_encerrados %>%
  mutate(ENCERRAMENTO_CAT = case_when(ENCERRAMENTO=="Interrupção do tratamento" ~ "Abandono",
                                      ENCERRAMENTO=="Óbito" ~ "Óbito",
                                      ENCERRAMENTO=="Tratamento completo" ~ "Tratamento completo",
                                      ENCERRAMENTO %in% c("Suspenso por condição clínica desfavorável ao tratamento","Suspenso por PT < 5mm em quimioprofilaxia primária",
                                                          "Suspenso por reação adversa")~"Tratamento suspenso",
                                      ENCERRAMENTO %in% c("Transferido para outro país","Tratamento completo","Tuberculose ativa")~ "Outros motivos",
                                      TRUE ~ "Indeterminado"
  )
  )
table(ILTB_nao_encerrados$ENCERRAMENTO_CAT)
tabela_encerramento <- prop.table(table(ILTB_nao_encerrados$ENCERRAMENTO_CAT)) * 100
print(tabela_encerramento)
table(ILTB_nao_encerrados$ENCERRAMENTO_CAT,ILTB_nao_encerrados$ano)
Encerramento_porcentagem <- prop.table(table(ILTB_nao_encerrados$ENCERRAMENTO_CAT,ILTB_nao_encerrados$ano), margin = 2) * 100
Encerramento_porcentagem

table(ILTB_2018_2024$Medicamentos)
#recategorizar medicamentos

ILTB_2018_2024 <- ILTB_2018_2024 %>%
  mutate(
    Medicamentos_cat = case_when(
      Medicamentos %in% c("Isoniazida", "Isoniazida - 9H", "Isoniazida - 6H") ~ "Isoniazida",
      Medicamentos %in% c("Rifampicina - 4R") ~ "Rifampicina",
      Medicamentos %in% c("Rifampicina + Isoniazida - 3RH (dispersíveis pediátricos)") ~ "Rifampicina_Isoniazida",
      Medicamentos %in% c("Rifapentina + Isoniazida - 3HP") ~ "Rifapentina_Isoniazida",
      TRUE ~ "Indeterminado"
    )
  )
table(ILTB_2018_2024$Medicamentos_cat)
MM_porcentagem<- prop.table(table(ILTB_2018_2024$Medicamentos_cat))*100
MM_porcentagem
table(ILTB_2018_2024$Medicamentos_cat,ILTB_2018_2024$ano)
MM_porcentagem_ano <- prop.table(table(ILTB_2018_2024$Medicamentos_cat, ILTB_2018_2024$ano), margin = 2) * 100
MM_porcentagem_ano



#RADIOGRAFIA DE TORAX: recategorizar 
ILTB_2018_2024 <- ILTB_2018_2024 %>% 
  mutate(
    Radiografia_cat = case_when(
      `Radiografia do Torax` %in% c("Alteração não sugestiva de TB ativa", "Alteração sugestiva de TB ativa", "Normal") ~ "Realizada",
      `Radiografia do Torax` == "Não realizada" ~ "Não realizada",
      TRUE ~ "Indeterminada"
    )
  )

table(ILTB_2018_2024$Radiografia_cat)
Radiografia_porcentagem<- prop.table(table(ILTB_2018_2024$Radiografia_cat))*100
Radiografia_porcentagem
table(ILTB_2018_2024$Radiografia_cat,ILTB_2018_2024$ano)
Radiografia_porcentagem_ano <- prop.table(table(ILTB_2018_2024$Radiografia_cat, ILTB_2018_2024$ano), margin = 2) * 100
Radiografia_porcentagem_ano


table(ILTB_2018_2024$IGRA)
#recategorizar
ILTB_2018_2024 <- ILTB_2018_2024 %>% 
  mutate(
    IGRA_cat = case_when(
      IGRA %in% c("Indeterminado", "Negativo", "Positivo") ~ "Realizado",
      IGRA == "Não realizado" ~ "Não Realizado",
      TRUE ~ "Indeterminado"
    )
  )
table(ILTB_2018_2024$IGRA_cat)
IGRA_porcentagem<- prop.table(table(ILTB_2018_2024$IGRA_cat))*100
IGRA_porcentagem
table(ILTB_2018_2024$IGRA_cat,ILTB_2018_2024$ano)
IGRA_porcentagem_ano <- prop.table(table(ILTB_2018_2024$IGRA_cat, ILTB_2018_2024$ano), margin = 2) * 100
IGRA_porcentagem_ano


table(ILTB_2018_2024$`Prova Tuberculinica (PT)`)
PPD_porcentagem<- prop.table(table(ILTB_2018_2024$`Prova Tuberculinica (PT)`))*100
PPD_porcentagem
table(ILTB_2018_2024$`Prova Tuberculinica (PT)`,ILTB_2018_2024$ano)
PPD_porcentagem_ano <- prop.table(table(ILTB_2018_2024$`Prova Tuberculinica (PT)`, ILTB_2018_2024$ano), margin = 2) * 100
PPD_porcentagem_ano

table(ILTB_2018_2024$HIV)
#recategorizar 
ILTB_2018_2024<-ILTB_2018_2024 %>%
  mutate(HIV_cat =
           case_when( HIV  %in% c("Em andamento","Negativo","Positivo") ~ "Realizado",
                                                               HIV =="Não realizado" ~ "Não realizado",
                                                               TRUE ~ "Indeterminado"
                      )
         )
table(ILTB_2018_2024$HIV_cat)
HIV_porcentagem<- prop.table(table(ILTB_2018_2024$HIV_cat))*100
HIV_porcentagem
table(ILTB_2018_2024$HIV_cat,ILTB_2018_2024$ano)
HIV_porcentagem_ano <- prop.table(table(ILTB_2018_2024$HIV_cat, ILTB_2018_2024$ano), margin = 2) * 100
HIV_porcentagem_ano

table(ILTB_2018_2024$`Tipo de entrada`)
Entrada_porcentagem<- prop.table(table(ILTB_2018_2024$`Tipo de entrada`))*100
Entrada_porcentagem
table(ILTB_2018_2024$`Tipo de entrada`,ILTB_2018_2024$ano)
Entrada_porcentagem_ano <- prop.table(table(ILTB_2018_2024$`Tipo de entrada`, ILTB_2018_2024$ano), margin = 2) * 100
Entrada_porcentagem_ano

table(ILTB_2018_2024$`Descartado TB ativa`)
TB_porcentagem<- prop.table(table(ILTB_2018_2024$`Descartado TB ativa`))*100
TB_porcentagem
table(ILTB_2018_2024$`Descartado TB ativa`,ILTB_2018_2024$ano)
TB_porcentagem_ano <- prop.table(table(ILTB_2018_2024$`Descartado TB ativa`, ILTB_2018_2024$ano), margin = 2) * 100
TB_porcentagem_ano




#GRÁFICO PIRÂMIDE ETÁRIA SEXO, FAIXA ETÁRIA E RAÇA/COR:

#Excluir categoria raça_cor ignorada
ILTB_grafico<- subset(ILTB_2018_2024,raca_cor_cat!="IGNORADO")
ILTB_grafico$População<-1 #para contar o numero de linhas

dados <- ILTB_grafico %>%
  mutate(População = ifelse(Sexo == "Masculino", -População, População))
# para colocar os dados referentes ao masculino do lado esquerdo da pirâmide


library(ggplot2)

ggplot(dados, aes(x = IDADEcat2, y = População, fill = raca_cor_cat)) +
  geom_bar(stat = "identity", position = "stack") +
  coord_flip() +
  geom_hline(yintercept = 0, linetype = "dashed", color = "black", size = 1.1) +  # Linha mais grossa
  scale_y_continuous(labels = abs) +
  labs(
    title = "Pirâmide Demográfica",
    x = "Faixa Etária",
    y = "População",
    fill = "Raça/cor"
  ) +
  theme_minimal() +
  theme(
    axis.text.x = element_text(size = 10),
    axis.text.y = element_text(size = 10)
  ) +
  scale_fill_brewer(palette = "Set1")


#Conferindo o gráfico:
table(ILTB_grafico$Sexo,ILTB_grafico$IDADEcat2)
sexo_idade_porcentagem_ano <- prop.table(table(ILTB_grafico$Sexo,ILTB_grafico$IDADEcat2), margin = 2) * 100
sexo_idade_porcentagem_ano

#saber se existe diferença entre sexo e faixa etária:
tabela <- table(ILTB_grafico$IDADEcat2, ILTB_grafico$Sexo)
print(tabela)
teste_qui2 <- chisq.test(tabela)
print(teste_qui2)
library(ggplot2)
ggplot(ILTB_grafico, aes(x = IDADEcat2, fill = Sexo)) +
  geom_bar(position = "fill") +
  labs(
    x = "Faixa Etária",
    y = "Proporção",
    fill = "Sexo",
    title = "Distribuição de Sexo por Faixa Etária"
  ) +
  theme_minimal()


table(ILTB_grafico$Sexo,ILTB_grafico$raca_cor_cat)
sexo_raca_cor_porcentagem_ano <- prop.table(table(ILTB_grafico$Sexo,ILTB_grafico$raca_cor), margin = 2) * 100
sexo_raca_cor_porcentagem_ano

#teste qui-quadrado
tabela2 <- table(ILTB_grafico$IDADEcat2, ILTB_grafico$raca_cor_cat)
print(tabela2)
teste_qui2 <- chisq.test(tabela2)
print(teste_qui2)
library(ggplot2)
ggplot(ILTB_grafico, aes(x = IDADEcat2, fill = Sexo)) +
  geom_bar(position = "fill") +
  labs(
    x = "Faixa Etária",
    y = "Proporção",
    fill = "Sexo",
    title = "Distribuição de Sexo por Faixa Etária"
  ) +
  theme_minimal()

table(ILTB_grafico$Sexo,ILTB_grafico$raca_cor_cat,ILTB_grafico$IDADEcat2)
sexo_raca_cor_idade_porcentagem <- prop.table(table(ILTB_grafico$Sexo,ILTB_grafico$raca_cor,ILTB_grafico$IDADEcat2), margin = 3) * 100
sexo_raca_cor_idade_porcentagem



##############   Usuários ativos site IL_TB e instituições    ###############

#USUARIOS<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Usuários_ILTB.xlsx")
#INSTITUICOES<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Instituições_ILTB.xlsx")
#MUNICIPIOS<-read.xlsx("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/MUNICIPIOS_POP.xlsx")

#Preparo dos bancos:

USUARIOS_ILTB<-subset(USUARIOS,Situação=="Ativo")
USUARIOS_ILTB <- USUARIOS_ILTB %>%
  group_by(Município) %>%
  mutate(Usuario_tot = n())
USUARIOS_ILTB <- USUARIOS_ILTB %>%
  group_by(Município) %>%
  mutate(Transacoes_tot = sum(`Transações nos últimos 12 meses`, na.rm = TRUE)) %>%
  ungroup()
USUARIOS_ILTB<-USUARIOS_ILTB[,c("Município","Usuario_tot","Transacoes_tot")]
USUARIOS_ILTB<-unique(USUARIOS_ILTB) #631 municípios possuem pelo menos 1 usuário ativo

INSTITUICOES_ILTB<-subset(INSTITUICOES,Situação=="Ativo")
INSTITUICOES_ILTB<-subset(INSTITUICOES_ILTB,`Endereço - Estado`=="Minas Gerais")
INSTITUICOES_ILTB<-INSTITUICOES_ILTB %>%
  group_by(`Endereço - Município`) %>%
  mutate(Instituicoes_tot = n())
names(INSTITUICOES_ILTB)[names(INSTITUICOES_ILTB)=="Endereço - Município"]<- "Município"
INSTITUICOES_ILTB<-INSTITUICOES_ILTB[,c("Município","Instituicoes_tot")]
INSTITUICOES_ILTB<-unique(INSTITUICOES_ILTB) #826 municípios possuem pelo menos 1 instituição cadastrada

#linkagem determinística dos bancos
USUARIOS_INSTITUICAO_ILTB <- merge(INSTITUICOES_ILTB, USUARIOS_ILTB, by = "Município", all = TRUE)
sum(is.na(USUARIOS_INSTITUICAO_ILTB$Instituicoes_tot)) # 0 obs
sum(is.na(USUARIOS_INSTITUICAO_ILTB$Transacoes_tot)) # 195 obs
sum(is.na(USUARIOS_INSTITUICAO_ILTB$Usuario_tot)) # 195
length(unique(USUARIOS_INSTITUICAO_ILTB$Município)) #826
#Existem 195 municípios com instituição cadastrada mas sem usuário ativo no ILTB


# COLOCAR NÚMERO ZERO NOS USUÁRIOS E NAS TRANSAÇÕES SEM INFORMAÇÃO
USUARIOS_INSTITUICAO_ILTB <- USUARIOS_INSTITUICAO_ILTB %>%
  mutate(across(everything(), ~ ifelse(is.na(.), 0, .)))
#Correção dos nomes dos municípios cujo nome do município no ILTB não corresponde ao nome no PDR
USUARIOS_INSTITUICAO_ILTB$Município<-gsub("Brazópolis","Brasópolis",USUARIOS_INSTITUICAO_ILTB$Município)
USUARIOS_INSTITUICAO_ILTB$Município<-gsub("Dona Eusébia","Dona Euzébia",USUARIOS_INSTITUICAO_ILTB$Município)
USUARIOS_INSTITUICAO_ILTB$Município<-gsub("Gouveia","Gouvêa",USUARIOS_INSTITUICAO_ILTB$Município)
USUARIOS_INSTITUICAO_ILTB$Município<-gsub("Olhos-d'Agua","Olhos-D'água",USUARIOS_INSTITUICAO_ILTB$Município)
USUARIOS_INSTITUICAO_ILTB$Município<-gsub("Pingo-d'Agua","Pingo-D'água",USUARIOS_INSTITUICAO_ILTB$Município)
USUARIOS_INSTITUICAO_ILTB$Município<-gsub("Queluzito","Queluzita",USUARIOS_INSTITUICAO_ILTB$Município)
USUARIOS_INSTITUICAO_ILTB$Município<-gsub("São João del Rei","São João Del Rei",USUARIOS_INSTITUICAO_ILTB$Município)
USUARIOS_INSTITUICAO_ILTB$Município<-gsub("São Pedro da UniAo","São Pedro da União",USUARIOS_INSTITUICAO_ILTB$Município)


#LINKAGEM COM O PDR PARA TRAZER A INFORMAÇÃO SOBRE URS:
MUNICIPIOS<-MUNICIPIOS[,c("NM_MUNICIP","Unidade.Regional.de.Saúde","POP_Municipio")]  
MUNICIPIOS$Unidade.Regional.de.Saúde<-gsub("Diamantina","DIAMANTINA",MUNICIPIOS$Unidade.Regional.de.Saúde)
names(MUNICIPIOS)[names(MUNICIPIOS)=="NM_MUNICIP"]<- "Município"
USUARIOS_INSTITUICAO_ILTB <- merge(MUNICIPIOS, USUARIOS_INSTITUICAO_ILTB, by = "Município", all = TRUE) #854 não vou sofrer!existe duas instituições sem município excluir depois Parear depois usando o código tbm
sum(is.na(USUARIOS_INSTITUICAO_ILTB$Instituicoes_tot)) # 28 obs
sum(is.na(USUARIOS_INSTITUICAO_ILTB$Unidade.Regional.de.Saúde)) #1 
# Existem 28 municípios que não possuem Instituição cadastrada

# COLOCAR NÚMERO ZERO NOS USUÁRIOS E NAS TRANSAÇÕES SEM INFORMAÇÃO
USUARIOS_INSTITUICAO_ILTB <- USUARIOS_INSTITUICAO_ILTB %>%
  mutate(across(everything(), ~ ifelse(is.na(.), 0, .)))

table(USUARIOS_INSTITUICAO_ILTB$Unidade.Regional.de.Saúde,USUARIOS_INSTITUICAO_ILTB$Instituicoes_tot)

soma_instituicao <- USUARIOS_INSTITUICAO_ILTB %>%
  group_by(Unidade.Regional.de.Saúde) %>%
  summarise(Soma_instituicao = sum(Instituicoes_tot, na.rm = TRUE))

soma_usuários <- USUARIOS_INSTITUICAO_ILTB %>%
  group_by(Unidade.Regional.de.Saúde) %>%
  summarise(Soma_usuarios = sum(Usuario_tot, na.rm = TRUE))

soma_transações <- USUARIOS_INSTITUICAO_ILTB %>%
  group_by(Unidade.Regional.de.Saúde) %>%
  summarise(Soma_transações = sum(Transacoes_tot, na.rm = TRUE))

Regionais_sem_instituicao<-subset(USUARIOS_INSTITUICAO_ILTB,Instituicoes_tot==0) #28 municípios
table(Regionais_sem_instituicao$Unidade.Regional.de.Saúde)

Regionais_sem_usuarios<-subset(USUARIOS_INSTITUICAO_ILTB,Usuario_tot==0) #223 municípios
table(Regionais_sem_usuarios$Unidade.Regional.de.Saúde)

Regionais_sem_transações<-subset(USUARIOS_INSTITUICAO_ILTB,Transacoes_tot==0) 
table(Regionais_sem_transações$Unidade.Regional.de.Saúde)

#write.xlsx(USUARIOS_INSTITUICAO_ILTB,"C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/IL-TB/Boletins/USUARIOS_INSTITUICAO_ILTB.xlsx") 




####################### RELAÇÃO ENTRE PRESENÇA DE HIV E TRATAMENTO NÃO ENCERRADO  

#Análise do desempenho das regionais:


####################################   END CODE ##############################################################################
#####################################28 DE FEVEREIRO DE 2025##################################################################
################################### THANK YOU MY LORD, THANK YOU MY MOTHER####################################################

