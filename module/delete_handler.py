def on_delete_button_clicked(main_window):
    currentIndex = main_window.file_combo.currentIndex()
    if currentIndex >= 0:
        main_window.file_combo.currentIndexChanged.disconnect()

        del main_window.file_list[currentIndex]
        main_window.file_combo.removeItem(currentIndex)

        if hasattr(main_window, 'header_table'):
            main_window.header_table.deleteLater()
            del main_window.header_table

        if len(main_window.file_list) > 0:
            newIndex = 0 if currentIndex == 0 else currentIndex - 1
            main_window.file_combo.setCurrentIndex(newIndex)
            main_window.on_file_combo_currentIndexChanged()

        main_window.file_combo.currentIndexChanged.connect(main_window.on_file_combo_currentIndexChanged)

    main_window.selected_columns = []
