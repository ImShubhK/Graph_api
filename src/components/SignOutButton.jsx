import Button from '@mui/material/Button';
import {useMsal} from "@azure/msal-react"

export const SignOutButton = () => {
    const {instance } =useMsal();
    const handelSignOut = ()=>{
        instance.logoutRedirect();
    }
    return (
        <Button color="inherit" onClick={handelSignOut}>Sign out</Button>
    )
};