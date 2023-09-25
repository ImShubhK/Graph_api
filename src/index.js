import React, { useEffect, useState } from 'react';
import ReactDOM from 'react-dom';
import { BrowserRouter } from 'react-router-dom';
import { ThemeProvider } from '@mui/material/styles';
import { theme } from './styles/theme';
import { PublicClientApplication, EventType, InteractionType } from '@azure/msal-browser';
import App from './App';

// Initialize MSAL instance
const pca = new PublicClientApplication({
    auth: {
        clientId: '936f51af-f57d-46c5-8461-40508707c9fc',
        authority: 'https://login.microsoftonline.com/common',
        redirectUri: '/',
    },
    cache: {
        cacheLocation: 'localStorage',
        storeAuthStateInCookie: true,
    },
});

const fetchCalendarEvents = async (accessToken) => {
    try {
        const response = await fetch('https://graph.microsoft.com/v1.0/me/events', {
            headers: {
                Authorization: `Bearer ${accessToken}`,
            },
        });

        if (response.ok) {
            const data = await response.json();
            return data.value; // Assuming the events are in the "value" property
        } else {
            console.error('Failed to fetch calendar events');
            throw new Error('Failed to fetch calendar events');
        }
    } catch (error) {
        console.error('Error fetching calendar events:', error);
        throw error;
    }
};
const fetchOutlookEmails = async (accessToken) => {
    try {
        const response = await fetch('https://graph.microsoft.com/v1.0/me/messages?$select=id,subject,body,sender,toRecipients', {
            headers: {
                Authorization: `Bearer ${accessToken}`,
            },
        });

        if (response.ok) {
            const data = await response.json();
            return data.value; // Assuming the emails are in the "value" property
        } else {
            console.error('Failed to fetch Outlook emails');
            throw new Error('Failed to fetch Outlook emails');
        }
    } catch (error) {
        console.error('Error fetching Outlook emails:', error);
        throw error;
    }
};

function AppWrapper() {
    const [calendarEvents, setCalendarEvents] = useState([]);
    const [outlookEmails, setOutlookEmails] = useState([]);

    useEffect(() => {
        pca.addEventCallback((event) => {
            if (event.eventType === EventType.LOGIN_SUCCESS) {
                console.log(event)
                // Extract the access token from the event's payload
                const accessToken = event.payload.accessToken;
    
                // Use the access token to fetch calendar events
                fetchCalendarEvents(accessToken)
                    .then((events) => {
                        // Log the fetched calendar events
                        console.log('Fetched Calendar Events:', events);
    
                        // Set the fetched calendar events in the component state
                        setCalendarEvents(events);
                    })
                    .catch((error) => {
                        // Handle any errors
                        console.error('Error fetching calendar events:', error);
                    });
    
                // Use the access token to fetch Outlook emails
                fetchOutlookEmails(accessToken)
                    .then((emails) => {
                        // Log the fetched Outlook emails
                        console.log('Fetched Outlook Emails:', emails);
    
                        // Set the fetched Outlook emails in the component state
                        setOutlookEmails(emails);
                    })
                    .catch((error) => {
                        // Handle any errors
                        console.error('Error fetching Outlook emails:', error);
                    });
            }
        });
    }, []);
    

    return (
        <React.StrictMode>
            <BrowserRouter>
                <ThemeProvider theme={theme}>
                    <App msalInstance={pca} />
                </ThemeProvider>
            </BrowserRouter>
            {/* Display calendar events or use them in your app */}
            <div>
    <h2>Calendar Events</h2>
    {calendarEvents.map((event) => (
        <div key={event.id}>
            <h3>{event.subject}</h3>
            {/* <p>{event.body?.content}</p> */}
            <p>Organizer: {event.organizer?.emailAddress?.name} ({event.organizer?.emailAddress?.address})</p>
            <p>Attendees:</p>
            <ul>
                {event.attendees?.map((attendee) => (
                    <li key={attendee.emailAddress?.address}>
                        {attendee.emailAddress?.name} ({attendee.emailAddress?.address})
                    </li>
                ))}
            </ul>
        </div>
    ))}
</div>

<div>
    <h2>Outlook Emails</h2>
    {outlookEmails.map((email) => (
        <div key={email.id}>
            <h3>{email.subject}</h3>
            <p>{email.body?.content}</p>
            <p>Sender: {email.sender?.emailAddress?.name} ({email.sender?.emailAddress?.address})</p>
            <p>Recipients:</p>
            <ul>
                {email.toRecipients?.map((recipient) => (
                    <li key={recipient.emailAddress?.address}>
                        {recipient.emailAddress?.name} ({recipient.emailAddress?.address})
                    </li>
                ))}
            </ul>
        </div>
    ))}
</div>

        </React.StrictMode>
    );
}

const root = ReactDOM.createRoot(document.getElementById('root'));
root.render(<AppWrapper />);
