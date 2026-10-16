
#include <iostream>
#include <xlsxwriter.h>

int main() {
    // Membuat file Excel
    lxw_workbook *workbook =
        workbook_new("hasil.xlsx");

    if (workbook == nullptr) {
        std::cerr << "Gagal membuat file Excel!\n";
        return 1;
    }

    // Membuat lembar kerja
    lxw_worksheet *worksheet =
        workbook_add_worksheet(workbook, "Data");

    // Membuat format judul
    lxw_format *header =
        workbook_add_format(workbook);

    format_set_bold(header);
    format_set_align(header, LXW_ALIGN_CENTER);

    // Mengatur lebar kolom
    worksheet_set_column(
        worksheet, 0, 0, 20, nullptr
    );
    worksheet_set_column(
        worksheet, 1, 1, 15, nullptr
    );
    worksheet_set_column(
        worksheet, 2, 2, 15, nullptr
    );

    // Membuat judul kolom
    worksheet_write_string(
        worksheet, 0, 0, "Nama", header
    );
    worksheet_write_string(
        worksheet, 0, 1, "Kelas", header
    );
    worksheet_write_string(
        worksheet, 0, 2, "Nilai", header
    );

    // Mengisi data
    worksheet_write_string(
        worksheet, 1, 0, "Andi", nullptr
    );
    worksheet_write_string(
        worksheet, 1, 1, "9A", nullptr
    );
    worksheet_write_number(
        worksheet, 1, 2, 90, nullptr
    );

    worksheet_write_string(
        worksheet, 2, 0, "Budi", nullptr
    );
    worksheet_write_string(
        worksheet, 2, 1, "9B", nullptr
    );
    worksheet_write_number(
        worksheet, 2, 2, 85, nullptr
    );

    worksheet_write_string(
        worksheet, 3, 0, "Citra", nullptr
    );
    worksheet_write_string(
        worksheet, 3, 1, "9C", nullptr
    );
    worksheet_write_number(
        worksheet, 3, 2, 95, nullptr
    );

    // Menyimpan dan menutup file Excel
    lxw_error result = workbook_close(workbook);

    if (result != LXW_NO_ERROR) {
        std::cerr << "Gagal menyimpan Excel!\n";
        return 1;
    }

    std::cout << "Berhasil membuat hasil.xlsx!\n";

    return 0;
}
