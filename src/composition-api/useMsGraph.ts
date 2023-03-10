import { injectStrict } from '@/utils/typescript-utils';
import { AxiosKey } from '@/types/symbols';

const baseUrl = 'https://graph.microsoft.com/v1.0/me';

export function useMsGraph() {
  const axios = injectStrict(AxiosKey);

  async function getGraphProfile() {
    try {
      const { data } = await axios.get<microsoftgraph.User>(baseUrl);
      return data;
    } catch (error) {
      console.log('Error getting user profile: ', error);
    }
  }

  async function getDriveFiles() {
    try {
      const { data } = await axios.get<microsoftgraph.DriveItem>(`${baseUrl}/drive/root/children`);
      return data;
    } catch (error) {
      console.log('Error getting drive items: ', error);
    }
  }

  async function getExcel(id: string) {
    try {
      const { data } = await axios.get<microsoftgraph.WorkbookWorksheet>(
        `${baseUrl}/drive/items/${id}/workbook/worksheets`,
      );
      return data;
    } catch (error) {
      console.log('Error getting worksheets: ', error);
    }
  }

  async function getTables(id: string, worksheetID: string) {
    try {
      const { data } = await axios.get<microsoftgraph.WorkbookTable>(
        `${baseUrl}/drive/items/${id}/workbook/worksheets/${worksheetID}/tables`,
      );
      return data;
    } catch (error) {
      console.log('Error getting tables: ', error);
    }
  }

  async function getColumns(fileID: string, worksheetID: string, tableID: string) {
    try {
      const { data } = await axios.get<microsoftgraph.WorkbookTableColumn>(
        `${baseUrl}/drive/items/${fileID}/workbook/worksheets/${worksheetID}/tables/${tableID}/columns`,
      );
      return data;
    } catch (error) {
      console.log('Error getting worksheets: ', error);
    }
  }

  async function postRow(fileID: string, worksheetID: string, tableID: string, payload: object) {
    try {
      const payloadObj = {
        values: [payload],
      };
      const { data } = await axios.post<microsoftgraph.WorkbookTableRow>(
        `${baseUrl}/drive/items/${fileID}/workbook/worksheets/${worksheetID}/tables/${tableID}/rows`,
        payloadObj,
      );

      return data;
    } catch (error) {
      console.log('Error creating row: ', error);
    }
  }

  return {
    getGraphProfile,
    getDriveFiles,
    getExcel,
    getTables,
    getColumns,
    postRow,
  };
}
