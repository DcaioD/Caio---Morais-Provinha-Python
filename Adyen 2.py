import pandas as pd
#import matplotlib.pyplot as plt
import os

# File data abse
df = pd.read_csv('DataAnalyst_case_study_data.csv')

# coluna amount ajustar 
df['amount'] = pd.to_numeric(df['amount'].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)

# KPI Geral
resumo_moeda = df.groupby('currency_code').agg(
    received=('psp_reference', 'count'),
    authorization=('authorization', 'sum')
).reset_index()
resumo_moeda['cancelled'] = resumo_moeda['received'] - resumo_moeda['authorization']
resumo_moeda['approved_rate'] = (resumo_moeda['authorization'] / resumo_moeda['received'] * 100).round(2).astype(str) + '%'
resumo_moeda['cancelled_rate'] = (resumo_moeda['cancelled'] / resumo_moeda['received'] * 100).round(2).astype(str) + '%'
resumo_moeda.columns = ['Currency ', 'received', 'authorization', 'Cancelled', 'Approved Rate', 'Cancelled Rate']
#print "Verificador de codigo 1 / erro no codigo"

# Babco pareto 
analise_bancos = df.groupby('issuername').agg(
    total_transacoes=('psp_reference', 'count'),
    autorizadas=('authorization', 'sum'))
analise_bancos['recusadas'] = analise_bancos['total_transacoes'] - analise_bancos['autorizadas']
analise_bancos['taxa_autorizacao'] = (analise_bancos['autorizadas'] / analise_bancos['total_transacoes']).round(4)
analise_bancos = analise_bancos.sort_values(by='recusadas', ascending=False)
total_rec_geral = analise_bancos['recusadas'].sum()
analise_bancos['pareto_acumulado_%'] = (analise_bancos['recusadas'] / total_rec_geral * 100).cumsum().round(2)
top_12_bancos = analise_bancos.head(12)
df_filtrado = df[df['issuername'].isin(top_12_bancos.index)].copy()
 #print "Verificador de codigo 2 / erro no codigo"

# Maiores erros
analise_motivos = df_filtrado.groupby('raw_acquirer_response').agg(
    total_transacoes=('psp_reference', 'count'),
    autorizadas=('authorization', 'sum') 
)
analise_motivos['recusadas'] = analise_motivos['total_transacoes'] - analise_motivos['autorizadas']
analise_motivos = analise_motivos[analise_motivos['recusadas'] > 0].sort_values(by='recusadas', ascending=False)
total_rec_filt = analise_motivos['recusadas'].sum()
analise_motivos['pareto_acumulado_%'] = (analise_motivos['recusadas'] / total_rec_filt * 100).cumsum().round(2)
top_10_motivos = analise_motivos[['total_transacoes', 'recusadas', 'pareto_acumulado_%']].head(10)
#print "Verificador de codigo 3 / erro no codigo"

# Print terminal para vizualixzar os erros / corrigor
print("\n KPI Geral ")
print(resumo_moeda.to_string(index=False))
print("\n Pareto Banco")
print(top_12_bancos.to_string())
print("\n Pareto Erros")
print(top_10_motivos.to_string())


# Faixa valor
df['faixa_valor'] = pd.cut(df['amount'], bins=[0, 50, 100, 500, 1000, 5000, 100000], 
                           labels=['0-50', '50-100', '100-500', '500-1000', '1000-5000', '5000+'])
analise_valor = df.groupby('faixa_valor', observed=False).agg(
    taxa_aprovacao=('authorization', 'mean'),
    contagem=('psp_reference', 'count')
).reset_index()


print("\n" + "="*90)
print("Analise de dados")
print("="*90)

 #print "Verificador de codigo 5 / erro no codigo"
top_5_bancos_nomes = top_12_bancos.head(5).index
df_top5_bancos = df[df['issuername'].isin(top_5_bancos_nomes)]

print("\n--- Taxa de aprovação (%) Por banco e valores")
tabela_bancos_valor = pd.crosstab(
    index=df_top5_bancos['issuername'], 
    columns=df_top5_bancos['faixa_valor'], 
    values=df_top5_bancos['authorization'], 
    aggfunc='mean'
).round(4) * 100
print(tabela_bancos_valor.fillna(0).to_string())

print("\n--- Volume de erro por banco")
df_recusadas_top5 = df_top5_bancos[df_top5_bancos['authorization'] == False]
erros_genericos = [m for m in df['raw_acquirer_response'].unique() if isinstance(m, str) and any(x in m for x in ['05', '62', '06'])]
df_erros_gen = df_recusadas_top5[df_recusadas_top5['raw_acquirer_response'].isin(erros_genericos)]

tabela_bancos_erros = pd.crosstab(
    index=df_erros_gen['issuername'],
    columns=df_erros_gen['raw_acquirer_response'],
    margins=True,
    margins_name="TOTAL_GERAL"
)
print(tabela_bancos_erros.to_string())
 #print "Verificador de codigo 6 / erro no codigo"


# Data export xlsx
print("\n" + "="*90)
print("Detalhamento dos top 4 erros/raw")
print("="*90)

excel_file = 'Relatorio_Adyen.xlsx'
writer = pd.ExcelWriter(excel_file, engine='openpyxl')

resumo_moeda.to_excel(writer, sheet_name='Resumo_Geral', index=False)
top_12_bancos.to_excel(writer, sheet_name='Pareto_Bancos')
top_10_motivos.to_excel(writer, sheet_name='Pareto_Motivos')
tabela_bancos_valor.to_excel(writer, sheet_name='Bancos_x_FaixaValor')
tabela_bancos_erros.to_excel(writer, sheet_name='Bancos_x_ErrosGenericos')

top_4_nomes = top_10_motivos.head(4).index
detalhes_para_excel = []

for i, motivo in enumerate(top_4_nomes, 1):
    df_motivo = df_filtrado[df_filtrado['raw_acquirer_response'] == motivo]
    
    if any(err in str(motivo) for err in ["51", "62", "57"]):
        colunas_agrupamento = ['shopper_interaction']
    else:
        colunas_agrupamento = ['shopper_interaction', 'cvc_data_supplied']

    colunas_presentes = [col for col in colunas_agrupamento if col in df_motivo.columns]

    if colunas_presentes:
        tabela_final = df_motivo.groupby(colunas_presentes).agg(
            qtd_recusas=('authorization', lambda x: (x == False).sum()),
            ticket_medio=('amount', 'mean')
        ).round(2).reset_index()
        
        # Imprimindo a tabela do detalhamento no terminal
        print(f"\nDETALHAMENTO {i}: {motivo}")
        print(tabela_final.to_string(index=False))

        # Verifcar algum ponto manualmente 
        try:
            if 'shopper_interaction' in tabela_final.columns:
                # Alerta Ecommerce vs ContAuth
                tk_cont_series = tabela_final[tabela_final['shopper_interaction'] == 'ContAuth']['ticket_medio']
                tk_ecom_series = tabela_final[tabela_final['shopper_interaction'] == 'Ecommerce']['ticket_medio']

                if not tk_cont_series.empty and not tk_ecom_series.empty:
                    tk_cont = tk_cont_series.mean()
                    tk_ecom = tk_ecom_series.mean()

                    if tk_ecom < tk_cont:
                        queda = tk_cont - tk_ecom
                        print(f"\n Analise ticket medio:")
                        print(f"    Ticket Médio ContAuth: R$ {tk_cont:.2f}")
                        print(f"    Ticket Médio Ecommerce: R$ {tk_ecom:.2f}")

        except Exception as e:
            pass

        # Alerta CVC
        if "05" in str(motivo) and 'cvc_data_supplied' in tabela_final.columns:
            try:
                ecom_no = tabela_final[(tabela_final['shopper_interaction'] == 'Ecommerce') & (tabela_final['cvc_data_supplied'] == 'No')]['ticket_medio'].values
                ecom_yes = tabela_final[(tabela_final['shopper_interaction'] == 'Ecommerce') & (tabela_final['cvc_data_supplied'] == 'Yes')]['ticket_medio'].values
                
                if len(ecom_no) > 0 and len(ecom_yes) > 0:
                    tk_ecom_no = ecom_no[0]
                    tk_ecom_yes = ecom_yes[0]
                    dif_cvc = tk_ecom_yes - tk_ecom_no
                    
                    print(f"\n Analise ticket medio:")
                    print(f"    Ticket Médio SEM CVC: R$ {tk_ecom_no:.2f}")
                    print(f"    Ticket Médio COM CVC: R$ {tk_ecom_yes:.2f}")
            except Exception as e:
                pass

        tabela_excel = tabela_final.copy()
        tabela_excel.insert(0, 'Motivo_Recusa', motivo)
        detalhes_para_excel.append(tabela_excel)

#
#fig1, ax1 = plt.subplots(figsize=(10, 6))
#ax1.bar(top_12_bancos.index, top_12_bancos['recusadas'], color='skyblue')
#ax1.set_title('Top 12 Bancos por Volume de Recusa', fontweight='bold')
#ax1.set_xticklabels(top_12_bancos.index, rotation=25, ha='right')
#ax2 = ax1.twinx()
#ax2.plot(top_12_bancos.index, top_12_bancos['pareto_acumulado_%'], color='red', marker='D')
#ax2.axhline(80, color='orange', linestyle='--')
#ax2.set_ylim(0, 105)
#fig2, ax3 = plt.subplots(figsize=(10, 6))
#ax3.bar(top_10_motivos.index, top_10_motivos['recusadas'], color='lightgreen')
#ax3.set_title('Top 10 Motivos de Recusa (Filtro: 12 Maiores Bancos)', fontweight='bold')
#ax3.set_xticklabels(top_10_motivos.index, rotation=25, ha='right')
#ax4 = ax3.twinx()
#ax4.plot(top_10_motivos.index, top_10_motivos['pareto_acumulado_%'], color='red', marker='D')
#ax4.set_ylim(0, 105)

if detalhes_para_excel:
    pd.concat(detalhes_para_excel, ignore_index=True).to_excel(writer, sheet_name='Detalhamento_Top4', index=False)

writer.close()
print(f"\n Excel  gerado e salvo: {excel_file}")

print("\nExcell com o relatorio do terminal foi aberto diretamente para analise")
os.system(f"open '{excel_file}'")