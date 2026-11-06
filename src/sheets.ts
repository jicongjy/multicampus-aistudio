import { RegistrationRecord } from './firebase';

export const SHEET_TAB_NAME = '참가자명단';

export const SHEET_HEADERS = [
  'ID',
  '신청 일시',
  '이름',
  '휴대폰 번호',
  '이메일',
  '참석 인원',
  '개인정보 동의',
  '소식 수신',
  '전달 사항',
  '수정 일시',
];

export function recordToRow(r: RegistrationRecord): (string | number)[] {
  return [
    r.id,
    r.createdAt ? new Date(r.createdAt).toLocaleString('ko-KR') : '',
    r.name,
    r.phone,
    r.email,
    r.count,
    r.agree ? '동의' : '미동의',
    r.mkt ? '동의' : '미동의',
    r.memo || '',
    r.updatedAt ? new Date(r.updatedAt).toLocaleString('ko-KR') : '',
  ];
}

/**
 * Create a new Google Spreadsheet in the user's Google Drive
 */
export async function createEventSpreadsheet(
  accessToken: string,
  title?: string
): Promise<{ spreadsheetId: string; spreadsheetUrl: string; title: string }> {
  const nowStr = new Date().toISOString().slice(0, 10);
  const sheetTitle = title || `이벤트 참가 신청 명단 (${nowStr})`;

  const res = await fetch('https://sheets.googleapis.com/v4/spreadsheets', {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      properties: {
        title: sheetTitle,
      },
      sheets: [
        {
          properties: {
            title: SHEET_TAB_NAME,
            gridProperties: {
              frozenRowCount: 1,
            },
          },
        },
      ],
    }),
  });

  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.error?.message || `스프레드시트 생성 실패 (상태 코드: ${res.status})`);
  }

  const data = await res.json();
  const spreadsheetId = data.spreadsheetId;
  const spreadsheetUrl = data.spreadsheetUrl || `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`;

  // Write headers to row 1
  await updateSheetRange(accessToken, spreadsheetId, `${SHEET_TAB_NAME}!A1:J1`, [SHEET_HEADERS]);

  // Format header row (bold text, dark slate background, white text)
  try {
    const sheetId = data.sheets?.[0]?.properties?.sheetId || 0;
    await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}:batchUpdate`, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        requests: [
          {
            repeatCell: {
              range: {
                sheetId: sheetId,
                startRowIndex: 0,
                endRowIndex: 1,
                startColumnIndex: 0,
                endColumnIndex: SHEET_HEADERS.length,
              },
              cell: {
                userEnteredFormat: {
                  backgroundColor: { red: 0.08, green: 0.12, blue: 0.18 }, // #141f2e
                  textFormat: {
                    bold: true,
                    foregroundColor: { red: 1.0, green: 1.0, blue: 1.0 },
                    fontSize: 10,
                  },
                  horizontalAlignment: 'CENTER',
                },
              },
              fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
            },
          },
        ],
      }),
    });
  } catch (e) {
    console.warn('Header formatting skipped:', e);
  }

  return { spreadsheetId, spreadsheetUrl, title: sheetTitle };
}

/**
 * Fetch spreadsheet metadata to check accessibility
 */
export async function getSpreadsheetInfo(
  accessToken: string,
  spreadsheetId: string
): Promise<{ title: string; tabExists: boolean }> {
  const res = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?fields=properties.title,sheets.properties.title`, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
    },
  });

  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.error?.message || '스프레드시트를 조회할 수 없습니다. 권한 또는 ID를 확인해 주세요.');
  }

  const data = await res.json();
  const sheets: any[] = data.sheets || [];
  const tabExists = sheets.some((s: any) => s.properties?.title === SHEET_TAB_NAME);

  return {
    title: data.properties?.title || '구글 스프레드시트',
    tabExists,
  };
}

/**
 * Update range with 2D values
 */
export async function updateSheetRange(
  accessToken: string,
  spreadsheetId: string,
  range: string,
  values: (string | number)[][]
): Promise<void> {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(range)}?valueInputOption=USER_ENTERED`;
  const res = await fetch(url, {
    method: 'PUT',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({ values }),
  });

  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.error?.message || `시트 업데이트 실패 (${res.status})`);
  }
}

/**
 * Clear existing data rows (from row 2 downwards)
 */
export async function clearSheetDataRows(accessToken: string, spreadsheetId: string): Promise<void> {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(`${SHEET_TAB_NAME}!A2:J`)}:clear`;
  const res = await fetch(url, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
    },
  });

  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.error?.message || `시트 데이터 초기화 실패 (${res.status})`);
  }
}

/**
 * Sync all registrations into the spreadsheet
 */
export async function syncAllRegistrationsToSpreadsheet(
  accessToken: string,
  spreadsheetId: string,
  records: RegistrationRecord[]
): Promise<number> {
  // Ensure header is present
  await updateSheetRange(accessToken, spreadsheetId, `${SHEET_TAB_NAME}!A1:J1`, [SHEET_HEADERS]);

  // Clear previous data rows
  await clearSheetDataRows(accessToken, spreadsheetId);

  if (records.length === 0) {
    return 0;
  }

  // Sort by ID ascending for spreadsheet readability
  const sorted = [...records].sort((a, b) => a.id - b.id);
  const rows = sorted.map(recordToRow);

  const range = `${SHEET_TAB_NAME}!A2:J${rows.length + 1}`;
  await updateSheetRange(accessToken, spreadsheetId, range, rows);
  return rows.length;
}

/**
 * Append single new registration to the spreadsheet
 */
export async function appendRegistrationToSpreadsheet(
  accessToken: string,
  spreadsheetId: string,
  record: RegistrationRecord
): Promise<void> {
  const row = recordToRow(record);
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(`${SHEET_TAB_NAME}!A:J`)}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`;

  const res = await fetch(url, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      values: [row],
    }),
  });

  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.error?.message || `스프레드시트 행 추가 실패 (${res.status})`);
  }
}
