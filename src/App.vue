<template>
  <div>
    <NavBar />
    <!--
    <img alt="Vue logo" src="./assets/logo.png">
    -->
    <h3 class="title">Microsoft 365 x Graph API x Vue3 App POC</h3>
    <router-view />
  </div>
</template>

<script lang="ts">
import { defineComponent } from 'vue';
import NavBar from "./components/NavBar.vue";

export default defineComponent({
  name: 'App',
  components: {
    NavBar
  }
});
</script>

<style>
#app {
  font-family: Avenir, Helvetica, Arial, sans-serif;
  text-align: center;
  color: #2c3e50;
}
.title {
  margin: 3vw;
}
</style>
