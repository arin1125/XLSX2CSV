def on_all_delete_button_clicked(main_window):
    main_window.file_combo.currentIndexChanged.disconnect()

    main_window.file_list.clear()
    main_window.file_combo.clear()

    if hasattr(main_window, 'header_table'):
        main_window.header_table.deleteLater()
        del main_window.header_table

    main_window.file_combo.currentIndexChanged.connect(main_window.on_file_combo_currentIndexChanged)

    main_window.selected_columns = []
