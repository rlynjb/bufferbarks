<template>
  <div class="home">
    <span v-if="!isAuthenticated">Please sign-in to see your profile information.</span>

    <div class="content" v-if="isAuthenticated">
      <div class="content-item">
        <el-button type="primary" v-on:click="goToProfile">Request Profile Information</el-button>
      </div>

      <div class="content-item resources">
        <h4>Resources</h4>
        <a
          href="https://learn.microsoft.com/en-us/graph/"
          target="_blank"
        >
          Documentation
        </a>
        <a
          href="https://learn.microsoft.com/en-us/graph/auth/"
          target="_blank"
        >Auth and Register App Documentation</a>
        <a
          href=""
          target="_blank"
        >Azure Portal</a>
        <a
          href="https://github.com/AzureAD/microsoft-authentication-library-for-js/tree/dev/samples/msal-browser-samples/vue3-sample-app"
          target="_blank">
          Vue3 Sample App with MSAL Authentication
        </a>
        <a
          href="https://developer.microsoft.com/en-us/graph/graph-explorer"
          target="_blank"
        >
          Graph Explorer
        </a>
      </div>

      <div class="content-item">
        <p>To work on Microsoft Apps, we need the root path of the file</p>
        <p>Look into setting permissions for: a single user, a group of users, global</p>
        <el-button @click="getDriveListItems">get files in my drive</el-button>

        <ul>
          <li v-for="file in state.driveListItems" :key="file.id">
            {{ file.name }}
            <el-button @click="getExcel(file.id)">
              Get workbook (excel) {{ file.id }}
            </el-button>
          </li>
        </ul>
      </div>

      <div class="content-item">
        <h4>
          Submit form and Add a Row in Excel Spreadsheet
        </h4>
        <input v-model="state.excelInput" />
        <el-button v-on:click="postNewExcelRow">submit</el-button>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { useRouter } from "vue-router";
import { useIsAuthenticated } from "../composition-api/useIsAuthenticated";

const isAuthenticated = useIsAuthenticated();

const router = useRouter();

// start - ms stuff
import { useMsal } from "../composition-api/useMsal";
import { InteractionRequiredAuthError, InteractionStatus } from "@azure/msal-browser";
import { loginRequest } from "../authConfig";
import {
  getGraphDriveItems, getGraphExcel, postGraphExcelRow
} from "../utils/MsGraphApiCall";

const { instance, inProgress } = useMsal();
// end - ms stuff

import { reactive } from "vue";

const state = reactive({
  excelInput: "",
  driveListItems: [],
});

async function getToken() {
  return await instance.acquireTokenSilent({
    ...loginRequest
  }).catch(async (e) => {
    if (e instanceof InteractionRequiredAuthError) {
      await instance.acquireTokenRedirect(loginRequest);
    }
    throw e;
  });
}

function goToProfile() {
  router.push("/profile");
}

async function getDriveListItems() {
  const response = await getToken();

  if (inProgress.value === InteractionStatus.None) {
    const graphData = await getGraphDriveItems(response.accessToken);
    state.driveListItems = graphData.value;
    console.log(state.driveListItems)
    //state.resolved = true;
    //stopWatcher();
  }
}

async function getExcel(id) {
	const response = await getToken();

	if (inProgress.value === InteractionStatus.None) {
		const graphData = await getGraphExcel(response.accessToken, id);
    console.log(graphData)
    //state.driveListItems = graphData.value;
		//state.resolved = true;
		//stopWatcher();
	}
}

async function postNewExcelRow() {
	const response = await getToken();

	if (inProgress.value === InteractionStatus.None) {
		const graphData = await postGraphExcelRow(response.accessToken);
		//state.data = graphData;
		//state.resolved = true;
		//stopWatcher();
	}
}
</script>

<style>
body {
  margin: 0;
}
h1, h2, h3, h4, h5, h6 {
  margin: 0;
}
.content {
  padding: 0vw 3vw 3vw;
}
.content-item {
  border: 1px solid #000;
  padding: 2vw;
  margin-bottom: 2vw;
}
.content-item.resources a {
  color: #000;
  margin: 2vw 2vw 0;
  display: inline-block;
}
</style>
