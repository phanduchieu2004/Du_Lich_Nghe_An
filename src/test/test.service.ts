import {
  BadRequestException,
  Injectable,
  NotFoundException,
} from '@nestjs/common';
import ExcelJS from 'exceljs';
import { Pool, type PoolClient } from 'pg';

export interface ColumnMetadata {
  columnName: string;
  dataType: string;
  isNullable: 'YES' | 'NO';
  columnDefault: string | null;
  isIdentity: 'YES' | 'NO';
  isGenerated: string;
  isPrimaryKey: boolean;
  udtName?: string;
}

interface TableMetadata {
  tableName: string;
  columns: ColumnMetadata[];
}

type JsonRecord = Record<string, unknown>;

const PAGE_SIZE = 50;

export interface TableRowsPage {
  table: string;
  schema: string;
  columns: ColumnMetadata[];
  primaryKey: string[];
  rows: JsonRecord[];
  total: number;
  page: number;
  pageSize: number;
}

interface PreparedImportRow {
  rowNumber: number;
  values: JsonRecord;
}

interface PreparedImportSheet {
  tableName: string;
  columns: ColumnMetadata[];
  rows: PreparedImportRow[];
}

interface PreparedWorkbook {
  sheets: PreparedImportSheet[];
  errors: string[];
  totalRows: number;
}

export interface ExcelImportResult {
  totalRows: number;
  tables: Array<{ tableName: string; rowCount: number }>;
}

export interface MediaLibraryItem {
  id: string;
  type: 'image' | 'audio' | 'video' | 'file';
  title: string;
  description: string | null;
  url: string;
  source: string;
  objectType: string | null;
  objectId: string | null;
}

const IMPORT_TABLE_ALIASES: Record<string, string> = {
  mau_1b_khu_vuc: 'phan_khu',
  '1b_khu_vuc': 'phan_khu',
  khu_vuc: 'phan_khu',
  mau_1c_hien_vat: 'hien_vat',
  '1c_hien_vat': 'hien_vat',
  mau_1d_tien_ich: 'tien_ich',
  '1d_tien_ich': 'tien_ich',
};

const IMPORT_COLUMN_ALIASES: Record<string, Record<string, string>> = {
  phan_khu: {
    name: 'ten_phan_khu',
    name_en: 'ten_tieng_anh',
    area_m2: 'dien_tich',
    mo_ta_chung: 'mo_ta',
  },
  hien_vat: {
    name: 'ten_hien_vat',
    name_en: 'ten_tieng_anh',
    tinh_trang_bao_quan: 'tinh_trang',
  },
  tien_ich: {
    name: 'ten_tien_ich',
    don_vi_van_hanh: 'don_vi_quan_ly',
    hotline_ho_tro: 'hotline',
    ghi_chu: 'mo_ta',
  },
};

const MAX_EXCEL_BYTES = 20 * 1024 * 1024;
const MAX_EXCEL_ROWS = 10_000;
const MAX_EXCEL_COLUMNS = 300;
const MAX_EXCEL_SHEETS = 100;
const MAX_IMPORT_ERRORS = 100;

@Injectable()
export class TestService {
  private pool: Pool | undefined;

  async getTables(): Promise<{ schema: string; tables: string[] }> {
    const schema = this.schemaName;
    const result = await this.getPool().query<{ tableName: string }>(
      `SELECT table_name AS "tableName"
       FROM information_schema.tables
       WHERE table_schema = $1 AND table_type = 'BASE TABLE'
         AND table_name <> 'spatial_ref_sys'
       ORDER BY table_name`,
      [schema],
    );

    return { schema, tables: result.rows.map(({ tableName }) => tableName) };
  }

  async getRows(
    tableName: string,
    search: string,
    page: number,
  ): Promise<TableRowsPage> {
    const table = await this.getTable(tableName);
    const primaryKey = table.columns
      .filter((column) => column.isPrimaryKey)
      .map((column) => column.columnName);
    const params: unknown[] = [];
    let where = '';
    const term = search.trim();

    if (term) {
      params.push(`%${term}%`);
      const searchConditions = table.columns.map(
        ({ columnName }) =>
          `${this.quoteIdentifier(columnName)}::text ILIKE $1`,
      );
      where = `WHERE (${searchConditions.join(' OR ')})`;
    }

    const qualifiedTable = this.qualifiedTable(table.tableName);
    const countResult = await this.getPool().query<{ count: string }>(
      `SELECT COUNT(*)::text AS count FROM ${qualifiedTable} ${where}`,
      params,
    );
    const total = Number(countResult.rows[0]?.count ?? 0);
    const offset = (page - 1) * PAGE_SIZE;
    const rowsResult = await this.getPool().query<JsonRecord>(
      `SELECT * FROM ${qualifiedTable} ${where}
       ORDER BY ${
         table.columns
           .filter((column) => column.isPrimaryKey)
           .map(({ columnName }) => this.quoteIdentifier(columnName))
           .join(', ') || '1'
       }
       LIMIT $${params.length + 1} OFFSET $${params.length + 2}`,
      [...params, PAGE_SIZE, offset],
    );

    return {
      table: table.tableName,
      schema: this.schemaName,
      columns: table.columns,
      primaryKey,
      rows: rowsResult.rows,
      total,
      page,
      pageSize: PAGE_SIZE,
    };
  }

  async getMediaLibrary(): Promise<MediaLibraryItem[]> {
    const schema = this.quoteIdentifier(this.schemaName);
    const pool = this.getPool();
    const [audio, video, media, mediaDanhNhan] = await Promise.all([
      pool.query(
        `SELECT ma_audio::text AS id, tieu_de AS title, NULL::text AS description,
                duong_dan_audio AS url, 'audio_thuyet_minh' AS source,
                loai_doi_tuong AS "objectType", ma_doi_tuong::text AS "objectId",
                'audio' AS "mediaType"
         FROM ${schema}."audio_thuyet_minh"
         WHERE NULLIF(BTRIM(duong_dan_audio), '') IS NOT NULL`,
      ),
      pool.query(
        `SELECT ma_video::text AS id, tieu_de AS title, mo_ta AS description,
                duong_dan_video AS url, 'video_thuyet_minh' AS source,
                loai_doi_tuong AS "objectType", ma_doi_tuong::text AS "objectId",
                'video' AS "mediaType"
         FROM ${schema}."video_thuyet_minh"
         WHERE NULLIF(BTRIM(duong_dan_video), '') IS NOT NULL`,
      ),
      pool.query(
        `SELECT media_id::text AS id, tieu_de AS title, mo_ta AS description,
                duong_dan AS url, 'media' AS source,
                loai_doi_tuong AS "objectType", ma_doi_tuong::text AS "objectId",
                loai_media AS "mediaType"
         FROM ${schema}."media"
         WHERE NULLIF(BTRIM(duong_dan), '') IS NOT NULL`,
      ),
      pool.query(
        `SELECT media_id::text AS id, 'Tư liệu danh nhân' AS title, mo_ta AS description,
                duong_dan AS url, 'media_danh_nhan' AS source,
                'danh_nhan' AS "objectType", ma_danh_nhan::text AS "objectId",
                loai_media AS "mediaType"
         FROM ${schema}."media_danh_nhan"
         WHERE NULLIF(BTRIM(duong_dan), '') IS NOT NULL`,
      ),
    ]);

    return [...audio.rows, ...video.rows, ...media.rows, ...mediaDanhNhan.rows]
      .filter((row) => this.isCloudinaryUrl(row.url))
      .map((row) => ({
        id: row.id as string,
        type: this.mediaType(row.mediaType, row.url as string),
        title:
          typeof row.title === 'string' && row.title.trim()
            ? row.title.trim()
            : 'Tư liệu thuyết minh',
        description:
          typeof row.description === 'string' ? row.description : null,
        url: row.url as string,
        source: row.source as string,
        objectType:
          typeof row.objectType === 'string' ? row.objectType : null,
        objectId: typeof row.objectId === 'string' ? row.objectId : null,
      }))
      .sort((first, second) => first.title.localeCompare(second.title, 'vi'));
  }

  async createRow(tableName: string, input: unknown): Promise<JsonRecord> {
    const table = await this.getTable(tableName);
    const body = this.readBody(input);
    const values = this.readValues(body.values, table.columns, false);
    const entries = Object.entries(values);
    const qualifiedTable = this.qualifiedTable(table.tableName);

    if (entries.length === 0) {
      const result = await this.getPool().query<JsonRecord>(
        `INSERT INTO ${qualifiedTable} DEFAULT VALUES RETURNING *`,
      );
      return result.rows[0];
    }

    const columnNames = entries.map(([name]) => this.quoteIdentifier(name));
    const placeholders = entries.map((_, index) => `$${index + 1}`);
    const result = await this.getPool().query<JsonRecord>(
      `INSERT INTO ${qualifiedTable} (${columnNames.join(', ')})
       VALUES (${placeholders.join(', ')}) RETURNING *`,
      entries.map(([, value]) => value),
    );

    return result.rows[0];
  }

  async updateRow(tableName: string, input: unknown): Promise<JsonRecord> {
    const table = await this.getTable(tableName);
    const primaryKey = this.requirePrimaryKey(table);
    const body = this.readBody(input);
    const key = this.readKey(body.key, primaryKey);
    const values = this.readValues(body.values, table.columns, true);
    const entries = Object.entries(values);
    if (entries.length === 0) {
      throw new BadRequestException(
        'Hãy nhập ít nhất một trường cần cập nhật.',
      );
    }

    const params = entries.map(([, value]) => value);
    const assignments = entries.map(
      ([name], index) => `${this.quoteIdentifier(name)} = $${index + 1}`,
    );
    const where = this.keyCondition(primaryKey, key, params.length + 1);
    const result = await this.getPool().query<JsonRecord>(
      `UPDATE ${this.qualifiedTable(table.tableName)}
       SET ${assignments.join(', ')} WHERE ${where.sql} RETURNING *`,
      [...params, ...where.values],
    );
    const row = result.rows[0];
    if (!row) {
      throw new NotFoundException('Không tìm thấy bản ghi cần cập nhật.');
    }
    return row;
  }

  async deleteRow(
    tableName: string,
    input: unknown,
  ): Promise<{ deleted: true }> {
    const table = await this.getTable(tableName);
    const primaryKey = this.requirePrimaryKey(table);
    const body = this.readBody(input);
    const key = this.readKey(body.key, primaryKey);
    const where = this.keyCondition(primaryKey, key, 1);
    const result = await this.getPool().query(
      `DELETE FROM ${this.qualifiedTable(table.tableName)}
       WHERE ${where.sql}`,
      where.values,
    );
    if (result.rowCount === 0) {
      throw new NotFoundException('Không tìm thấy bản ghi cần xóa.');
    }
    return { deleted: true };
  }

  async importExcel(buffer: Buffer): Promise<ExcelImportResult> {
    const prepared = await this.prepareExcelImport(buffer);
    if (prepared.errors.length > 0) {
      throw new BadRequestException(
        this.toImportError(prepared, prepared.errors),
      );
    }
    if (prepared.totalRows === 0) {
      throw new BadRequestException(
        this.toImportError(prepared, [
          'Không tìm thấy dòng dữ liệu nào để nhập.',
        ]),
      );
    }

    const client = await this.getPool().connect();
    try {
      await client.query('BEGIN');
      await client.query("SET LOCAL statement_timeout = '30s'");
      await this.insertPreparedRows(client, prepared);
      await this.syncImportSequences(client, prepared);
      await client.query('COMMIT');
      return this.toImportResult(prepared);
    } catch (error) {
      await client.query('ROLLBACK');
      throw new BadRequestException(
        {
          message: this.formatImportError(error),
          errors: [this.formatImportError(error)],
          ...this.toImportResult(prepared),
        },
      );
    } finally {
      client.release();
    }
  }

  private async getTable(tableName: string): Promise<TableMetadata> {
    const result = await this.getPool().query<ColumnMetadata>(
      `SELECT
         c.column_name AS "columnName",
         c.data_type AS "dataType",
         c.udt_name AS "udtName",
         c.is_nullable AS "isNullable",
         c.column_default AS "columnDefault",
         c.is_identity AS "isIdentity",
         c.is_generated AS "isGenerated",
         EXISTS (
           SELECT 1
           FROM information_schema.table_constraints tc
           JOIN information_schema.key_column_usage kcu
             ON kcu.constraint_catalog = tc.constraint_catalog
            AND kcu.constraint_schema = tc.constraint_schema
            AND kcu.constraint_name = tc.constraint_name
           WHERE tc.constraint_type = 'PRIMARY KEY'
             AND tc.table_schema = c.table_schema
             AND tc.table_name = c.table_name
             AND kcu.column_name = c.column_name
         ) AS "isPrimaryKey"
       FROM information_schema.columns c
       WHERE c.table_schema = $1 AND c.table_name = $2
       ORDER BY c.ordinal_position`,
      [this.schemaName, tableName],
    );

    if (result.rows.length === 0) {
      throw new NotFoundException(
        `Không tìm thấy bảng "${tableName}" trong schema.`,
      );
    }

    return { tableName, columns: result.rows };
  }

  private async prepareExcelImport(buffer: Buffer): Promise<PreparedWorkbook> {
    this.ensureTestMode();
    if (buffer.length === 0 || buffer.length > MAX_EXCEL_BYTES) {
      return {
        sheets: [],
        totalRows: 0,
        errors: ['File Excel trống hoặc vượt quá giới hạn 20 MB.'],
      };
    }

    const workbook = new ExcelJS.Workbook();
    try {
      const arrayBuffer = buffer.buffer.slice(
        buffer.byteOffset,
        buffer.byteOffset + buffer.byteLength,
      ) as ArrayBuffer;
      await workbook.xlsx.load(arrayBuffer);
    } catch {
      return {
        sheets: [],
        totalRows: 0,
        errors: ['Không đọc được file. Hãy chọn một file Excel .xlsx hợp lệ.'],
      };
    }

    if (workbook.worksheets.length === 0) {
      return {
        sheets: [],
        totalRows: 0,
        errors: ['File Excel phải có ít nhất một sheet dữ liệu.'],
      };
    }
    if (workbook.worksheets.length > MAX_EXCEL_SHEETS) {
      return {
        sheets: [],
        totalRows: 0,
        errors: [`File vượt quá giới hạn ${MAX_EXCEL_SHEETS} sheet.`],
      };
    }

    const { tables: knownTables } = await this.getTables();
    const columnResult = await this.getPool().query<{
      tableName: string;
      columnName: string;
    }>(
      `SELECT table_name AS "tableName", column_name AS "columnName"
       FROM information_schema.columns
       WHERE table_schema = $1 AND table_name = ANY($2::text[])`,
      [this.schemaName, knownTables],
    );
    const tableColumns = new Map<string, Set<string>>();
    for (const { tableName, columnName } of columnResult.rows) {
      const columns = tableColumns.get(tableName) ?? new Set<string>();
      columns.add(columnName.toLocaleLowerCase('en-US'));
      tableColumns.set(tableName, columns);
    }

    const sheets: PreparedImportSheet[] = [];
    const errors: string[] = [];
    const seenTargetTables = new Set<string>();
    let totalRows = 0;
    const addError = (error: string) => {
      if (errors.length < MAX_IMPORT_ERRORS) errors.push(error);
    };

    for (const worksheet of workbook.worksheets) {
      const headerRow = worksheet.getRow(1);
      if (!headerRow.hasValues) {
        addError(`Sheet "${worksheet.name}": thiếu dòng tiêu đề ở dòng 1.`);
        continue;
      }
      if (headerRow.cellCount > MAX_EXCEL_COLUMNS) {
        addError(
          `Sheet "${worksheet.name}": vượt quá giới hạn ${MAX_EXCEL_COLUMNS} cột.`,
        );
        continue;
      }

      const headerNames: string[] = [];
      for (let columnNumber = 1; columnNumber <= headerRow.cellCount; columnNumber += 1) {
        const value = headerRow.getCell(columnNumber).value;
        if (value === null || value === undefined || value === '') continue;
        if (typeof value !== 'string') {
          addError(
            `Sheet "${worksheet.name}", dòng 1, cột ${columnNumber}: tiêu đề cột phải là tên trường dạng chữ.`,
          );
          continue;
        }
        headerNames.push(value.trim());
      }

      const tableName = this.resolveImportTable(
        worksheet.name,
        headerNames,
        knownTables,
        tableColumns,
      );
      if (!tableName) {
        addError(
          `Sheet "${worksheet.name}": không xác định được bảng đích từ các cột. Đặt tên sheet trùng tên bảng, dùng tên sheet mẫu như "mau_1b_khu_vuc", hoặc bảo đảm có ít nhất hai cột tương ứng duy nhất.`,
        );
        continue;
      }
      if (seenTargetTables.has(tableName)) {
        addError(
          `Sheet "${worksheet.name}": nhiều sheet cùng được nhận diện là bảng "${tableName}". Gộp dữ liệu vào một sheet để tránh nhập trùng.`,
        );
        continue;
      }
      seenTargetTables.add(tableName);
      const table = await this.getTable(tableName);
      const databaseColumns = new Map(
        table.columns.map((column) => [
          column.columnName.toLocaleLowerCase('en-US'),
          column,
        ]),
      );
      const aliases = IMPORT_COLUMN_ALIASES[tableName] ?? {};
      const headerColumns = new Map<number, ColumnMetadata>();
      const sourceColumnNames = new Map<number, string>();
      const seenHeaders = new Set<string>();
      const seenTargetColumns = new Set<string>();

      for (
        let columnNumber = 1;
        columnNumber <= headerRow.cellCount;
        columnNumber += 1
      ) {
        const value = headerRow.getCell(columnNumber).value;
        if (value === null || value === undefined || value === '') continue;
        if (typeof value !== 'string') {
          continue;
        }

        const columnName = value.trim();
        const normalizedName = columnName.toLocaleLowerCase('en-US');
        if (seenHeaders.has(normalizedName)) {
          addError(
            `Sheet "${worksheet.name}": tiêu đề "${columnName}" bị lặp.`,
          );
          continue;
        }
        seenHeaders.add(normalizedName);
        const mappedColumnName =
          aliases[normalizedName]?.toLocaleLowerCase('en-US') ?? normalizedName;
        const column = databaseColumns.get(mappedColumnName);
        if (!column) {
          addError(
            `Sheet "${worksheet.name}", dòng 1: trường "${columnName}" không có cột tương ứng trong bảng "${tableName}".`,
          );
          continue;
        }
        if (seenTargetColumns.has(column.columnName)) {
          addError(
            `Sheet "${worksheet.name}": nhiều tiêu đề cùng ánh xạ đến trường "${column.columnName}" của bảng "${tableName}".`,
          );
          continue;
        }
        seenTargetColumns.add(column.columnName);
        headerColumns.set(columnNumber, column);
        sourceColumnNames.set(columnNumber, columnName);
      }

      if (headerColumns.size === 0) {
        addError(
          `Sheet "${worksheet.name}": không có trường nào tương ứng với các cột của bảng "${tableName}".`,
        );
        continue;
      }

      const rows: PreparedImportRow[] = [];
      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;
        if (!row.hasValues) return;
        if (totalRows >= MAX_EXCEL_ROWS) {
          addError(
            `File vượt quá giới hạn ${MAX_EXCEL_ROWS} dòng dữ liệu tổng cộng.`,
          );
          return;
        }
        if (row.cellCount > MAX_EXCEL_COLUMNS) {
          addError(
            `Sheet "${worksheet.name}", dòng ${rowNumber}: vượt quá giới hạn ${MAX_EXCEL_COLUMNS} cột.`,
          );
          return;
        }

        const values: JsonRecord = {};
        let hasValue = false;
        for (
          let columnNumber = 1;
          columnNumber <= row.cellCount;
          columnNumber += 1
        ) {
          const cell = row.getCell(columnNumber);
          const raw = cell.value;
          if (raw === null || raw === undefined || raw === '') continue;
          hasValue = true;
          const column = headerColumns.get(columnNumber);
          const sourceColumnName =
            sourceColumnNames.get(columnNumber) ?? String(columnNumber);
          if (!column) {
            addError(
              `Sheet "${worksheet.name}", dòng ${rowNumber}, cột ${columnNumber}: có dữ liệu nhưng không có tiêu đề CSDL tương ứng.`,
            );
            continue;
          }

          if (column.isGenerated !== 'NEVER') {
            addError(
              `Sheet "${worksheet.name}", dòng ${rowNumber}: cột "${sourceColumnName}" (${column.columnName}) do CSDL tự sinh, không được nhập giá trị.`,
            );
            continue;
          }
          if (raw && typeof raw === 'object' && 'formula' in raw) {
            addError(
              `Sheet "${worksheet.name}", dòng ${rowNumber}: không hỗ trợ công thức Excel tại cột "${sourceColumnName}".`,
            );
            continue;
          }

          try {
            values[column.columnName] = this.convertExcelCell(
              raw,
              column,
              workbook.properties.date1904,
            );
          } catch (error) {
            addError(
              `Sheet "${worksheet.name}", dòng ${rowNumber}, cột "${sourceColumnName}": ${error instanceof Error ? error.message : 'giá trị không hợp lệ.'}`,
            );
          }
        }

        if (!hasValue) return;
        for (const column of table.columns) {
          if (
            !Object.hasOwn(values, column.columnName) &&
            column.isNullable === 'NO' &&
            !column.columnDefault &&
            column.isIdentity !== 'YES' &&
            column.isGenerated === 'NEVER'
          ) {
            addError(
              `Sheet "${worksheet.name}", dòng ${rowNumber}: thiếu trường bắt buộc "${column.columnName}".`,
            );
          }
        }

        rows.push({ rowNumber, values });
        totalRows += 1;
      });

      sheets.push({ tableName, columns: table.columns, rows });
    }

    return { sheets, errors, totalRows };
  }

  private convertExcelCell(
    value: ExcelJS.CellValue,
    column: ColumnMetadata,
    date1904 = false,
  ): unknown {
    if (value === '\\N') return null;

    let cellValue: unknown = value;
    if (
      typeof value === 'number' &&
      (column.dataType === 'date' || column.dataType.startsWith('timestamp'))
    ) {
      const excelEpoch = date1904
        ? Date.UTC(1904, 0, 1)
        : Date.UTC(1899, 11, 30);
      cellValue = new Date(excelEpoch + value * 24 * 60 * 60 * 1000);
    }
    if (cellValue instanceof Date) {
      if (Number.isNaN(cellValue.getTime())) {
        throw new Error('ngày hoặc giờ Excel không hợp lệ.');
      }
      const isoValue = cellValue.toISOString();
      return column.dataType === 'date' ? isoValue.slice(0, 10) : isoValue;
    }
    if (typeof value === 'object' && value !== null) {
      if ('richText' in value && Array.isArray(value.richText)) {
        cellValue = value.richText.map(({ text }) => text).join('');
      } else if ('text' in value && typeof value.text === 'string') {
        cellValue = value.hyperlink ?? value.text;
      } else {
        throw new Error('kiểu ô Excel này không được hỗ trợ.');
      }
    }

    if (column.dataType === 'boolean') {
      if (typeof cellValue === 'boolean') return cellValue;
      if (typeof cellValue === 'string') {
        const normalized = cellValue.trim().toLocaleLowerCase('vi-VN');
        if (['true', '1', 'yes', 'có'].includes(normalized)) return true;
        if (['false', '0', 'no', 'không'].includes(normalized)) return false;
      }
      throw new Error('giá trị đúng/sai phải là Có/Không hoặc true/false.');
    }

    if (
      ['smallint', 'integer', 'bigint'].includes(column.udtName ?? '') &&
      typeof cellValue === 'number' &&
      !Number.isSafeInteger(cellValue)
    ) {
      throw new Error(
        'số nguyên vượt độ chính xác của Excel; hãy nhập ô này dưới dạng văn bản.',
      );
    }
    if (
      [
        'numeric',
        'decimal',
        'real',
        'float4',
        'float8',
        'double precision',
      ].includes(column.dataType) ||
      ['int2', 'int4', 'int8'].includes(column.udtName ?? '')
    ) {
      if (typeof cellValue !== 'number' && typeof cellValue !== 'string') {
        throw new Error('trường số phải là số hoặc chuỗi số.');
      }
      if (typeof cellValue === 'number' && !Number.isFinite(cellValue)) {
        throw new Error('trường số phải là một giá trị hữu hạn.');
      }
    }

    if (column.dataType === 'json' || column.dataType === 'jsonb') {
      if (typeof cellValue === 'string') {
        try {
          return JSON.parse(cellValue) as unknown;
        } catch {
          throw new Error('JSON không hợp lệ.');
        }
      }
    }

    return cellValue;
  }

  private async insertPreparedRows(
    client: PoolClient,
    prepared: PreparedWorkbook,
  ): Promise<void> {
    const populatedSheets = prepared.sheets.filter(
      (sheet) => sheet.rows.length > 0,
    );
    const order = await this.getImportOrder(
      client,
      populatedSheets.map(({ tableName }) => tableName),
    );
    const sheetsByName = new Map(
      populatedSheets.map((sheet) => [sheet.tableName, sheet]),
    );

    for (const tableName of order) {
      const sheet = sheetsByName.get(tableName);
      if (!sheet) continue;

      for (const row of sheet.rows) {
        const entries = Object.entries(row.values);
        const qualifiedTable = this.qualifiedTable(tableName);
        const columnNames = entries.map(([name]) => this.quoteIdentifier(name));
        const params = entries.map(([, value]) => value);
        const placeholders = entries.map((_, index) => `$${index + 1}`);
        const identityColumns = sheet.columns.some(
          (column) =>
            column.isIdentity === 'YES' &&
            Object.hasOwn(row.values, column.columnName),
        );
        const identityOverride = identityColumns
          ? ' OVERRIDING SYSTEM VALUE'
          : '';
        const statement =
          entries.length === 0
            ? `INSERT INTO ${qualifiedTable} DEFAULT VALUES`
            : `INSERT INTO ${qualifiedTable} (${columnNames.join(', ')})${identityOverride}
               VALUES (${placeholders.join(', ')})`;

        try {
          await client.query(statement, params);
        } catch (error) {
          throw new Error(
            this.formatImportError(error, tableName, row.rowNumber),
          );
        }
      }
    }
  }

  private async getImportOrder(
    client: PoolClient,
    tableNames: string[],
  ): Promise<string[]> {
    const included = new Set(tableNames);
    const indegree = new Map(tableNames.map((name) => [name, 0]));
    const children = new Map(
      tableNames.map((name) => [name, new Set<string>()]),
    );
    const result = await client.query<{
      parentTable: string;
      childTable: string;
    }>(
      `SELECT DISTINCT
         ccu.table_name AS "parentTable",
         tc.table_name AS "childTable"
       FROM information_schema.table_constraints tc
       JOIN information_schema.constraint_column_usage ccu
         ON ccu.constraint_catalog = tc.constraint_catalog
        AND ccu.constraint_schema = tc.constraint_schema
        AND ccu.constraint_name = tc.constraint_name
       WHERE tc.constraint_schema = $1 AND tc.constraint_type = 'FOREIGN KEY'`,
      [this.schemaName],
    );

    for (const { parentTable, childTable } of result.rows) {
      if (
        parentTable !== childTable &&
        included.has(parentTable) &&
        included.has(childTable) &&
        !children.get(parentTable)?.has(childTable)
      ) {
        children.get(parentTable)?.add(childTable);
        indegree.set(childTable, (indegree.get(childTable) ?? 0) + 1);
      }
    }

    const ready = tableNames.filter((name) => indegree.get(name) === 0);
    const order: string[] = [];
    while (ready.length > 0) {
      const table = ready.shift();
      if (!table) continue;
      order.push(table);
      for (const child of children.get(table) ?? []) {
        const remaining = (indegree.get(child) ?? 0) - 1;
        indegree.set(child, remaining);
        if (remaining === 0) ready.push(child);
      }
    }

    if (order.length !== tableNames.length) {
      const unresolved = tableNames.filter((name) => !order.includes(name));
      throw new Error(
        `Quan hệ khóa ngoại tạo vòng lặp giữa các sheet: ${unresolved.join(', ')}. Không thể xác định thứ tự nhập an toàn.`,
      );
    }
    return order;
  }

  private async syncImportSequences(
    client: PoolClient,
    prepared: PreparedWorkbook,
  ): Promise<void> {
    for (const sheet of prepared.sheets) {
      if (sheet.rows.length === 0) continue;
      for (const column of sheet.columns) {
        if (!column.columnDefault?.startsWith('nextval(')) continue;
        const sequenceResult = await client.query<{
          sequenceName: string | null;
        }>('SELECT pg_get_serial_sequence($1, $2) AS "sequenceName"', [
          `${this.schemaName}.${sheet.tableName}`,
          column.columnName,
        ]);
        const sequenceName = sequenceResult.rows[0]?.sequenceName;
        if (!sequenceName) continue;
        const tableName = this.qualifiedTable(sheet.tableName);
        const columnName = this.quoteIdentifier(column.columnName);
        await client.query(
          `SELECT setval(
             $1::regclass,
             GREATEST(
               COALESCE((SELECT MAX(${columnName}) FROM ${tableName}), 1),
               COALESCE(
                 (SELECT last_value FROM pg_sequences WHERE schemaname = $2 AND sequencename = $3),
                 1
               )
             ),
             true
           )`,
          [sequenceName, this.schemaName, sequenceName.split('.').pop()],
        );
      }
    }
  }

  private resolveImportTable(
    sheetName: string,
    headerNames: string[],
    knownTables: string[],
    tableColumns: Map<string, Set<string>>,
  ): string | undefined {
    const normalizedSheetName = sheetName.trim().toLocaleLowerCase('en-US');
    if (knownTables.includes(sheetName)) return sheetName;
    if (knownTables.includes(normalizedSheetName)) return normalizedSheetName;

    const aliasTable = IMPORT_TABLE_ALIASES[normalizedSheetName];
    if (aliasTable && knownTables.includes(aliasTable)) return aliasTable;

    const scores = knownTables.map((tableName) => {
      const columns = tableColumns.get(tableName) ?? new Set<string>();
      const aliases = IMPORT_COLUMN_ALIASES[tableName] ?? {};
      const matched = new Set<string>();
      for (const headerName of headerNames) {
        const normalizedHeader = headerName.toLocaleLowerCase('en-US');
        const targetColumn = aliases[normalizedHeader] ?? normalizedHeader;
        if (columns.has(targetColumn)) matched.add(targetColumn);
      }
      return { tableName, score: matched.size };
    });
    const bestScore = Math.max(0, ...scores.map(({ score }) => score));
    const bestMatches = scores.filter(({ score }) => score === bestScore);
    if (bestScore < 2 || bestMatches.length !== 1) return undefined;
    return bestMatches[0]?.tableName;
  }

  private toImportResult(prepared: PreparedWorkbook): ExcelImportResult {
    return {
      totalRows: prepared.totalRows,
      tables: prepared.sheets
        .filter((sheet) => sheet.rows.length > 0)
        .map(({ tableName, rows }) => ({ tableName, rowCount: rows.length })),
    };
  }

  private toImportError(prepared: PreparedWorkbook, errors: string[]) {
    const tables = this.toImportResult(prepared).tables;
    return {
      message: errors[0] ?? 'File Excel không hợp lệ.',
      errors: errors.slice(0, MAX_IMPORT_ERRORS),
      totalRows: prepared.totalRows,
      tables,
    };
  }

  private formatImportError(
    error: unknown,
    tableName?: string,
    rowNumber?: number,
  ): string {
    if (error instanceof Error && !('code' in error)) return error.message;
    const code =
      typeof error === 'object' && error !== null && 'code' in error
        ? String(error.code)
        : '';
    const reason: Record<string, string> = {
      '23503': 'khóa ngoại không khớp bản ghi cha.',
      '23505': 'bị trùng khóa chính hoặc giá trị duy nhất.',
      '23502': 'thiếu giá trị bắt buộc.',
      '23514': 'vi phạm điều kiện kiểm tra của CSDL.',
      '22P02': 'giá trị không đúng kiểu dữ liệu.',
      '22001': 'giá trị dài hơn giới hạn của cột.',
      '22003': 'giá trị số vượt giới hạn của cột.',
      '22007': 'ngày hoặc giờ không hợp lệ.',
    };
    const location =
      tableName && rowNumber ? `Bảng "${tableName}", dòng ${rowNumber}: ` : '';
    return `${location}${reason[code] ?? 'CSDL từ chối dữ liệu.'}`;
  }

  private ensureTestMode(): void {
    if (process.env.NODE_ENV === 'production') {
      throw new NotFoundException();
    }
  }

  private readValues(
    input: unknown,
    columns: ColumnMetadata[],
    updating: boolean,
  ): JsonRecord {
    const values = this.readBody(input);
    const metadata = new Map(
      columns.map((column) => [column.columnName, column]),
    );

    for (const name of Object.keys(values)) {
      const column = metadata.get(name);
      if (!column) {
        throw new BadRequestException(
          `Trường "${name}" không tồn tại trong bảng.`,
        );
      }
      if (column.isGenerated !== 'NEVER' || column.isIdentity === 'YES') {
        throw new BadRequestException(`Trường "${name}" được CSDL tự sinh.`);
      }
      if (updating && column.isPrimaryKey) {
        throw new BadRequestException('Không thể thay đổi khóa chính.');
      }
    }

    return values;
  }

  private readBody(input: unknown): JsonRecord {
    if (input === null || typeof input !== 'object' || Array.isArray(input)) {
      throw new BadRequestException(
        'Dữ liệu gửi lên phải là một đối tượng JSON.',
      );
    }
    return input as JsonRecord;
  }

  private readKey(input: unknown, primaryKey: string[]): JsonRecord {
    const key = this.readBody(input);
    const keyNames = Object.keys(key);
    if (
      keyNames.length !== primaryKey.length ||
      primaryKey.some((name) => !Object.hasOwn(key, name))
    ) {
      throw new BadRequestException('Cần gửi đủ giá trị của khóa chính.');
    }
    return key;
  }

  private requirePrimaryKey(table: TableMetadata): string[] {
    const columns = table.columns
      .filter((column) => column.isPrimaryKey)
      .map((column) => column.columnName);
    if (columns.length === 0) {
      throw new BadRequestException(
        'Bảng không có khóa chính nên không thể sửa hoặc xóa an toàn.',
      );
    }
    return columns;
  }

  private keyCondition(
    columns: string[],
    key: JsonRecord,
    firstParameter: number,
  ): { sql: string; values: unknown[] } {
    return {
      sql: columns
        .map(
          (column, index) =>
            `${this.quoteIdentifier(column)} = $${firstParameter + index}`,
        )
        .join(' AND '),
      values: columns.map((column) => key[column]),
    };
  }

  private qualifiedTable(tableName: string): string {
    return `${this.quoteIdentifier(this.schemaName)}.${this.quoteIdentifier(tableName)}`;
  }

  private quoteIdentifier(identifier: string): string {
    return `"${identifier.replaceAll('"', '""')}"`;
  }

  private isCloudinaryUrl(value: unknown): value is string {
    if (typeof value !== 'string') {
      return false;
    }
    try {
      const url = new URL(value);
      return (
        url.protocol === 'https:' &&
        (url.hostname === 'cloudinary.com' ||
          url.hostname.endsWith('.cloudinary.com'))
      );
    } catch {
      return false;
    }
  }

  private mediaType(value: unknown, mediaUrl: string): MediaLibraryItem['type'] {
    const path = new URL(mediaUrl).pathname.toLocaleLowerCase('en-US');
    const typeName = typeof value === 'string' ? value : '';
    const normalized = typeName
      .toLocaleLowerCase('vi')
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '');
    if (/(audio|am_thanh|am-thanh)/.test(normalized)) {
      return 'audio';
    }
    if (/(video)/.test(normalized)) {
      return 'video';
    }
    if (/(image|hinh|anh|photo)/.test(normalized)) {
      return 'image';
    }

    const extension = path.match(/\.([a-z0-9]+)$/)?.[1];
    if (
      extension &&
      ['mp3', 'wav', 'ogg', 'm4a', 'aac', 'flac', 'opus'].includes(extension)
    ) {
      return 'audio';
    }
    if (
      extension &&
      ['mp4', 'mov', 'm4v', 'webm', 'avi', 'mkv', 'm3u8'].includes(extension)
    ) {
      return 'video';
    }
    if (
      extension &&
      ['jpg', 'jpeg', 'png', 'webp', 'gif', 'avif', 'svg'].includes(extension)
    ) {
      return 'image';
    }

    if (path.includes('/image/upload/')) {
      return 'image';
    }
    if (path.includes('/video/upload/')) {
      return 'video';
    }
    return 'file';
  }

  private get schemaName(): string {
    return process.env.TEST_DB_SCHEMA || 'public';
  }

  private getPool(): Pool {
    this.ensureTestMode();
    this.pool ??= new Pool({
      host: process.env.DB_HOST || 'localhost',
      port: Number(process.env.DB_PORT) || 5432,
      user: process.env.DB_USERNAME || 'postgres',
      password: process.env.DB_PASSWORD,
      database: process.env.DB_DATABASE || 'du_lich',
    });
    return this.pool;
  }
}
