        b.metric("Vencidos", expired)
        c.metric("Vencen hoy", today_count)
        d.metric("Próximos a vencer", soon)
        e.metric("Fecha por revisar", unknown)

        st.markdown("#### Vencimientos por categoría")
        by_category = all_records.groupby("Categoria", dropna=False).size().reset_index(name="Cantidad").sort_values("Cantidad", ascending=False)
        st.dataframe(by_category, use_container_width=True, hide_index=True)

        st.markdown("#### Acciones prioritarias")
        priority = all_records[
            all_records["Dias_restantes"].apply(
                lambda x: pd.notna(x) and (x < 0 or x == 0 or (x > 0 and x <= max(warning_days or [30])))
            )
        ].sort_values("Dias_restantes", na_position="last")
        if priority.empty:
            st.success("No hay vencimientos prioritarios según la anticipación configurada.")
        else:
            st.dataframe(
                priority[["Categoria", "Tipo_control", "Proyecto", "Persona_activo", "Identificador",
                          "Fecha_vencimiento", "Dias_restantes", "Estado_seguimiento",
                          "Responsable", "Correo", "Archivo", "Hoja", "source_row"]],
                use_container_width=True, hide_index=True
            )
        st.caption("Los estados son cálculos administrativos basados en fechas. Confirma las obligaciones y periodicidades con el área responsable.")

with tab_sources:
    st.subheader("Archivos y hojas importados")
    if all_sources.empty:
        st.info("Todavía no se han importado archivos.")
    else:
        sources_view = all_sources.rename(columns={
            "source_file": "Archivo", "source_hash": "Huella del archivo",
            "uploaded_by": "Cargado por", "sheets": "Hoja",
            "imported_at": "Fecha de importación"
        })
        st.dataframe(sources_view[["Archivo", "Hoja", "Cargado por", "Fecha de importación", "Huella del archivo"]], use_container_width=True, hide_index=True)
        st.caption("La huella permite distinguir versiones del archivo. Al importar de nuevo una hoja con el mismo nombre, sus registros anteriores se reemplazan por la versión nueva.")

with tab_records:
    st.subheader("Buscar registros consolidados")
    if all_records.empty:
        st.info("No hay registros para mostrar.")
    else:
        f1, f2, f3 = st.columns(3)
        cat_options = options_for(all_records, "Categoria")
        selected_cat = f1.selectbox("Categoría", cat_options)
        project_options = options_for(all_records, "Proyecto")
        selected_project = f2.selectbox("Proyecto / área", project_options)
        search = f3.text_input("Buscar por persona, placa, código, archivo o responsable").strip().lower()
        view = all_records.copy()
        if selected_cat != "Todos":
            view = view[view["Categoria"].astype(str) == selected_cat]
        if selected_project != "Todos":
            view = view[view["Proyecto"].astype(str) == selected_project]
        if search:
            search_cols = ["Persona_activo", "Identificador", "Archivo", "Responsable", "Tipo_control", "Proyecto"]
            mask = False
            for col in search_cols:
                mask = mask | view[col].astype(str).str.lower().str.contains(search, na=False)
            view = view[mask]
        display = view[[
            "Categoria", "Tipo_control", "Proyecto", "Persona_activo", "Identificador",
            "Fecha_emision", "Fecha_vencimiento", "Dias_restantes", "Estado_seguimiento",
            "Responsable", "Correo", "Archivo", "Hoja", "source_row"
        ]].rename(columns={"source_row": "Fila origen"})
        st.dataframe(display, use_container_width=True, hide_index=True)

with tab_export:
    st.subheader("Exportar reportes")
    if all_records.empty:
        st.info("Importa al menos una hoja para habilitar reportes.")
    else:
        export_cols = [
            "Categoria", "Tipo_control", "Proyecto", "Persona_activo", "Identificador",
            "Fecha_emision", "Fecha_vencimiento", "Dias_restantes", "Estado_seguimiento",
            "Responsable", "Correo", "Archivo", "Hoja", "source_row", "uploaded_by"
        ]
        report = all_records[export_cols].rename(columns={
            "source_row": "Fila_origen", "uploaded_by": "Cargado_por"
        }).copy()
        priority_report = report[
            report["Estado_seguimiento"].isin(["Vencido", "Vence hoy", "Próximo a vencer"])
            | report["Estado_seguimiento"].str.startswith("Vence en", na=False)
        ]
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            report.to_excel(writer, index=False, sheet_name="Todos_los_vencimientos")
            priority_report.to_excel(writer, index=False, sheet_name="Acciones_prioritarias")
            if not all_sources.empty:
                all_sources.to_excel(writer, index=False, sheet_name="Archivos_importados")
        st.download_button(
            "⬇️ Descargar reporte corporativo Excel",
            data=output.getvalue(),
            file_name=f"reporte_corporativo_vencimientos_{today.isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

st.divider()
st.caption(
    "Prototipo administrativo. Antes de usar datos reales, valida con TI, SST y el responsable de privacidad "
    "los permisos, el acceso a la aplicación, la retención de archivos y el almacenamiento. "
    "No cargues diagnósticos ni historias clínicas. Los archivos originales se guardan en data/originales."
)# seguimiento-emo-applus
Sistema de seguimiento de exámenes médicos ocupacionales
