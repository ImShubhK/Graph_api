import Typography from "@mui/material/Typography";
import {useMsal}from "@azure/msal-react";
import { useEffect,useState } from "react";
export const WelcomeName = () => {

    const{instance}=useMsal();
    const[userName,setUserName ] =useState('');
    useEffect(()=>{
  const currentAccount = instance.getActiveAccount();
  if(currentAccount){
    setUserName(currentAccount.username);
  }
    },[instance])

    return <Typography variant="h6">Welcome, {userName}</Typography>;
};