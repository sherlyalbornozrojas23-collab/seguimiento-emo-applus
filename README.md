        report_filters_1, report_filters_2 = st.columns(2)
        report_cat_options = options_for(all_records, "Categoria")
        report_cat = report_filters_1.selectbox("Filtrar informe por categoría", report_cat_options, key="report_cat")
        report_project_options = options_for(all_records, "Proyecto")
        report_project = report_filters_2.selectbox("Filtrar informe por proyecto / área", report_project_options, key="report_project")
        report_data = all_records.copy()
        if report_cat != "Todos":
            report_data = report_data[report_data["Categoria"].astype(str) == report_cat]
        if report_project != "Todos":
            report_data = report_data[report_data["Proyecto"].astype(str) == report_project]

        status_summary = report_data.groupby("Estado_seguimiento").size().reset_index(name="Cantidad").sort_values("Cantidad", ascending=False)
        cat_summary = report_data.groupby("Categoria").size().reset_index(name="Cantidad").sort_values("Cantidad", ascending=False)
        project_summary = report_data.groupby("Proyecto").size().reset_index(name="Cantidad").sort_values("Cantidad", ascending=False)
        st.markdown("#### Vista gráfica del informe")
        report_chart_1, report_chart_2 = st.columns(2)
        with report_chart_1:
            fig_report_status = px.bar(status_summary, x="Estado_seguimiento", y="Cantidad", color="Estado_seguimiento", text="Cantidad", title="Resumen por estado")
            fig_report_status.update_layout(xaxis_title="Estado", yaxis_title="Cantidad", xaxis_tickangle=-25, height=360)
            fig_report_status.update_traces(textposition="outside")
            st.plotly_chart(fig_report_status, use_container_width=True, key="report_status_chart")
        with report_chart_2:
            fig_report_cat = px.bar(cat_summary, x="Categoria", y="Cantidad", color="Categoria", text="Cantidad", title="Resumen por categoría")
            fig_report_cat.update_layout(xaxis_title="Categoría", yaxis_title="Cantidad", xaxis_tickangle=-25, height=360, showlegend=False)
            fig_report_cat.update_traces(textposition="outside")
            st.plotly_chart(fig_report_cat, use_container_width=True, key="report_category_chart")

        export_cols = [
            "Categoria", "Tipo_control", "Proyecto", "Persona_activo", "Identificador",
            "Fecha_emision", "Fecha_vencimiento", "Dias_restantes", "Estado_seguimiento",
            "Responsable", "Correo", "Archivo", "Hoja", "source_row", "uploaded_by"
        ]
        report = report_data[export_cols].rename(columns={
            "source_row": "Fila_origen", "uploaded_by": "Cargado_por"
        }).copy()
        priority_report = report[
            report["Estado_seguimiento"].isin(["Vencido", "Vence hoy", "Próximo a vencer"])
            | report["Estado_seguimiento"].str.startswith("Vence en", na=False)
        ]
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            report.to_excel(writer, index=False, sheet_name="Base_consolidada")
            priority_report.to_excel(writer, index=False, sheet_name="Acciones_prioritarias")
            cat_summary.to_excel(writer, index=False, sheet_name="Resumen_categoria")
            status_summary.to_excel(writer, index=False, sheet_name="Resumen_estado")
            project_summary.to_excel(writer, index=False, sheet_name="Resumen_proyecto")
            if not all_sources.empty:
                all_sources.to_excel(writer, index=False, sheet_name="Archivos_importados")
            # Add bar charts to the Excel workbook so the report includes visual summaries.
            from openpyxl.chart import BarChart, Reference
            from openpyxl.styles import Font, PatternFill, Alignment
            workbook = writer.book
            for sheet_name, title in [("Resumen_categoria", "Vencimientos por categoría"), ("Resumen_estado", "Vencimientos por estado"), ("Resumen_proyecto", "Vencimientos por proyecto")]:
                ws = workbook[sheet_name]
                ws.freeze_panes = "A2"
                ws.auto_filter.ref = ws.dimensions
                for cell in ws[1]:
                    cell.font = Font(bold=True, color="FFFFFF")
                    cell.fill = PatternFill("solid", fgColor="174A7E")
                    cell.alignment = Alignment(horizontal="center")
                for col_cells in ws.columns:
                    letter = col_cells[0].column_letter
                    max_len = max((len(str(c.value or "")) for c in col_cells), default=10)
                    ws.column_dimensions[letter].width = min(max(max_len + 2, 12), 36)
                if ws.max_row > 1:
                    chart = BarChart()
                    chart.type = "bar"
                    chart.style = 10
                    chart.title = title
                    chart.y_axis.title = "Cantidad"
                    chart.x_axis.title = "Grupo"
                    data = Reference(ws, min_col=2, min_row=1, max_row=ws.max_row)
                    cats = Reference(ws, min_col=1, min_row=2, max_row=ws.max_row)
                    chart.add_data(data, titles_from_data=True)
                    chart.set_categories(cats)
                    chart.height = 7
                    chart.width = 15
                    ws.add_chart(chart, "D2")
            for sheet_name in ["Base_consolidada", "Acciones_prioritarias", "Archivos_importados"]:
                if sheet_name in workbook.sheetnames:
                    ws = workbook[sheet_name]
                    ws.freeze_panes = "A2"
                    ws.auto_filter.ref = ws.dimensions
                    for cell in ws[1]:
                        cell.font = Font(bold=True, color="FFFFFF")
                        cell.fill = PatternFill("solid", fgColor="174A7E")
                    for col_cells in ws.columns:
                        letter = col_cells[0].column_letter
                        max_len = max((len(str(c.value or "")) for c in col_cells[:100]), default=10)
                        ws.column_dimensions[letter].width = min(max(max_len + 2, 12), 36)

        st.download_button(
            "⬇️ Descargar informe ejecutivo Excel con gráficos",
            data=output.getvalue(),
            file_name=f"informe_corporativo_vencimientos_{today.isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.caption("El Excel incluye la base consolidada, acciones prioritarias, resúmenes y gráficos de barras. Los gráficos del panel web responden a los filtros; los del Excel reflejan los datos incluidos al descargarlo.")

st.divider()
st.caption(
    "Prototipo administrativo. Antes de usar datos reales, valida con TI, SST y el responsable de privacidad "
    "los permisos, el acceso a la aplicación, la retención de archivos y el almacenamiento. "
    "No cargues diagnósticos ni historias clínicas. Los archivos originales se guardan en data/originales."
)

st.caption(
    "Prototipo administrativo. Antes de usar datos reales, valida con TI, SST y el responsable de privacidad "
    "los permisos, el acceso a la aplicación, la retención de archivos y el almacenamiento. "
    "No cargues diagnósticos ni historias clínicas. Los archivos originales se guardan en data/originales."
)# seguimiento-emo-applus
Sistema de seguimiento de exámenes médicos ocupacionales
