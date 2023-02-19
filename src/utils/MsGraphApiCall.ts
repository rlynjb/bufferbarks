import { loginRequest } from "../authConfig";
import { useMsal } from "../composition-api/useMsal";
const { instance, inProgress } = useMsal();

const domain = 'https://graph.microsoft.com/v1.0/me';


export async function getGraphProfile(accessToken: string) {
    const headers = new Headers();
    headers.append("Authorization", `Bearer ${accessToken}`);

    const options = {
        method: "GET",
        headers: headers
    };

    return fetch(domain, options)
        .then(response => response.json())
        .catch(error => {
            console.log(error);
            throw error;
        });
}


export async function getGraphDriveItems(accessToken: string) {
    const headers = new Headers();
    headers.append("Authorization", `Bearer ${accessToken}`);

    const options = {
        method: "GET",
        headers: headers
    };

    return fetch(`${domain}/drive/root/children`, options)
        .then(response => response.json())
        .catch(error => {
            console.log(error);
            throw error;
        });
}

export async function getGraphExcel(accessToken: string, id: string) {
    const headers = new Headers();
    headers.append("Authorization", `Bearer ${accessToken}`);

    const options = {
        method: "GET",
        headers: headers
    };

    return fetch(`${domain}/drive/items/${id}/workbook/worksheets`, options)
        .then(response => response.json())
        .catch(error => {
            console.log(error);
            throw error;
        });
}

export async function postGraphExcelRow(accessToken: string) {
    const headers = new Headers();
    headers.append("Authorization", `Bearer ${accessToken}`);

    const payload = {
        "values": [
            [
                "alex darrow",
                "123",
                "adarrow@tenant.onmicrosoft.com"
            ]
        ]
    }

    const options = {
        method: "POST",
        headers: headers,
        body: JSON.stringify(payload)
    };

    return fetch(
        `${domain}/drive/items/01K6AWNMKR4C6YCVVJHZG3ZTWJVJDXQCN6/workbook/tables/Table1/rows`,
        options
    )
        .then(response => response.json())
        .catch(error => {
            console.log(error);
            throw error;
        });
}