import React, { useEffect, useState } from 'react';
import { app } from '@microsoft/teams-js';

interface TeamsContext {
  user?: { id?: string; displayName?: string };
  team?: { groupId?: string; displayName?: string };
  channelId?: string;
  tenantId?: string;
}

function App() {
  const [ctx, setCtx] = useState<TeamsContext>({});
  const [status, setStatus] = useState<string>('');
  const [msg, setMsg] = useState<string>('');
  const [sending, setSending] = useState<boolean>(false);

  useEffect(() => {
    app.initialize().then(() => {
      app.getContext().then((context) => {
        setCtx({
          user: {
            id: context.user?.id,
            displayName: context.user?.displayName,
          },
          team: {
            groupId: context.team?.groupId,
            displayName: context.team?.displayName,
          },
          channelId: context.channel?.id,
          tenantId: (context as any).tenant?.id || (context as any).tid,
        });
      });
    });
  }, []);

  const sendToBackend = async () => {
    if (!msg.trim()) {
      setStatus('Bitte eine Nachricht eingeben.');
      return;
    }
    setSending(true);
    setStatus('Sende an Backend...');
    try {
      const resp = await fetch('https://daad01dbd1f0.ngrok-free.app/aiagent', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          message: msg,
          teamsContext: ctx
        })
      });
      const data = await resp.json();
      if (resp.ok) setStatus('Backend-Antwort: ' + (data.reply || JSON.stringify(data)));
      else setStatus('Fehler beim Senden: ' + (data?.error || resp.status));
    } catch (err: any) {
      setStatus('Fetch-Fehler: ' + err.message);
    }
    setSending(false);
  };

  return (
    <div style={{ fontFamily: 'Arial', margin: 20 }}>
      <h2>Teams Tab Demo – Kontextdaten</h2>
      <pre>{JSON.stringify(ctx, null, 2)}</pre>
      <input
        value={msg}
        onChange={e => setMsg(e.target.value)}
        placeholder="Nachricht an den Agent..."
        style={{ width: '60%', marginRight: 8 }}
        disabled={sending}
      />
      <button onClick={sendToBackend} disabled={sending}>
        {sending ? 'Sendet...' : 'Daten ans Backend senden'}
      </button>
      <div style={{ marginTop: 8 }}>{status}</div>
    </div>
  );
}

export default App;
