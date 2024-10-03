import ExcelJS from "exceljs";

/**
 * Create an Excel file and return buffer or save to file
 * @param {Object} options - Configuration options
 * @param {Array<Object>} options.data - Array of data objects to export
 * @param {Array<Object>} options.columns - Array of column definitions
 * @param {string} [options.filename] - Output filename (optional if download is false)
 * @param {string} [options.sheetName='Sheet1'] - Worksheet name
 * @param {Object} [options.styles] - Custom styles for the Excel file
 * @param {boolean} [options.download=false] - If true, saves file and returns path. If false, returns buffer
 * @returns {Promise<Object>} Promise that resolves to { buffer, filename } or { path, filename }
 */

export const CREATE_EXCEL = async (options) => {
    try {
        // Validate required options
        if (!options.data || !Array.isArray(options.data)) {
            throw new Error('Data is required and must be an array');
        }
        if (!options.columns || !Array.isArray(options.columns)) {
            throw new Error('Columns configuration is required and must be an array');
        }
        
        const download = options.download ?? false;
        if (download && !options.filename) {
            throw new Error('Filename is required when download is true');
        }

        // Default values
        const defaultStyles = {
            headerFont: { bold: true, color: { argb: 'FF0000FF' } },
            headerFill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF00' } },
            border: {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            }
        };

        const styles = { ...defaultStyles, ...options.styles };
        const sheetName = options.sheetName || 'Sheet1';

        // Create workbook and worksheet
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet(sheetName);

        // Set columns
        worksheet.columns = options.columns;

        // Add filter
        worksheet.autoFilter = {
            from: { row: 1, column: 1 },
            to: { row: 1, column: options.columns.length }
        };

        // Style header row
        const headerRow = worksheet.getRow(1);
        headerRow.font = styles.headerFont;

        // Apply styles to header cells
        headerRow.eachCell((cell, colNumber) => {
            const column = options.columns[colNumber - 1];
            if (column && column.header) {
                cell.fill = styles.headerFill;
            }
            cell.border = styles.border;
        });

        // Add data rows
        options.data.forEach(item => {
            const row = worksheet.addRow(item);
            row.eachCell((cell) => {
                cell.border = styles.border;
            });
        });

        // Auto-size columns
        worksheet.columns.forEach(column => {
            let maxLength = 0;
            column.eachCell({ includeEmpty: true }, (cell) => {
                if (cell.row.number > 1 && cell.value) {
                    const length = cell.value.toString().length;
                    maxLength = Math.max(maxLength, length);
                }
            });
            column.width = maxLength < 10 ? 10 : maxLength;
        });

        if (download) {
            // Save to file and return path
            await workbook.xlsx.writeFile(options.filename);
            return {
                path: options.filename,
                filename: options.filename,
                success: true,
                message: `Excel file created: ${options.filename}`
            };
        } else {
            // Return buffer
            const buffer = await workbook.xlsx.writeBuffer();
            return {
                buffer,
                filename: options.filename || 'download.xlsx',
                success: true,
                message: 'Excel buffer created successfully'
            };
        }
    } catch (error) {
        console.error('Error creating Excel file:', error);
        throw error;
    }
};