if es_admin:
        with st.expander("📂 Sube o Actualiza tu Archivo de Clientes con Coordenadas"):
            st.markdown("Sube tu archivo `.xlsx` o `.csv` para actualizar el catálogo de clientes.")
            
            archivo_excel = st.file_uploader("Sube tu archivo de clientes", type=["xlsx", "xls", "csv"], key="uploader_masivo_clientes")
            
            if archivo_excel is not None:
                try:
                    if archivo_excel.name.endswith('.csv'):
                        df_preview = pd.read_csv(archivo_excel)
                    else:
                        df_preview = pd.read_excel(archivo_excel)
                        
                    if not df_preview.empty:
                        columnas_disponibles = list(df_preview.columns)
                        
                        def buscar_indice(palabras_clave):
                            for idx, col in enumerate(columnas_disponibles):
                                if any(p in str(col).lower() for p in palabras_clave):
                                    return idx
                            return 0

                        # Mapeo simplificado: Solo lo necesario (Cliente, Sucursal, Latitud, Longitud)
                        st.markdown("🗺️ **Mapeo Automático de Columnas:**")
                        c1, c2 = st.columns(2)
                        with c1:
                            col_cli = st.selectbox("Columna de Cliente / ID:", columnas_disponibles, index=buscar_indice(["cliente", "id", "num"]))
                            col_dir = st.selectbox("Columna de Sucursal / Nombre:", columnas_disponibles, index=buscar_indice(["nombre", "sucursal", "direccion"]))
                        with c2:
                            col_lat = st.selectbox("Columna de LATITUD:", columnas_disponibles, index=buscar_indice(["lat"]))
                            col_lon = st.selectbox("Columna de LONGITUD:", columnas_disponibles, index=buscar_indice(["lon"]))
                        
                        st.markdown("")
                        if st.button("📥 Procesar y Cargar Archivo"):
                            nueva_db = {}
                            lista_c = []
                            contador_registros = 0
                            
                            for _, row in df_preview.iterrows():
                                nc = str(row[col_cli]).strip() if pd.notna(row[col_cli]) else ""
                                nd = str(row[col_dir]).strip() if pd.notna(row[col_dir]) else "Principal"
                                
                                try: lat_val = float(row[col_lat]) if pd.notna(row[col_lat]) else 25.844412
                                except: lat_val = 25.844412
                                
                                try: lon_val = float(row[col_lon]) if pd.notna(row[col_lon]) else -100.395043
                                except: lon_val = -100.395043

                                if nc and nc.lower() != 'nan':
                                    if nc not in lista_c:
                                        lista_c.append(nc)
                                    if nc not in nueva_db:
                                        nueva_db[nc] = {}
                                    
                                    nueva_db[nc][nd] = {
                                        "latitud": lat_val,
                                        "longitud": asegurar_longitud_negativa(lon_val)
                                    }
                                    contador_registros += 1

                            if lista_c:
                                st.session_state.lista_clientes = lista_c
                                st.session_state.db_clientes = nueva_db
                                guardar_catalogos_en_disco()
                                
                                # Confirmación clara y detallada de que fue cargado con éxito
                                st.success(f"✅ ¡Archivo cargado y agregado con éxito! Se procesaron **{len(lista_c)} clientes** y **{contador_registros} sucursales/ubicaciones** correctamente.")
                except Exception as e:
                    st.error(f"Error al procesar el archivo: {e}")
