type EmailComposeParams = {
  to: string;
  subject: string;
  body?: string;
};

export function buildMailtoUrl({ to, subject, body }: EmailComposeParams) {
  const params = new URLSearchParams();
  params.set("subject", subject);
  if (body) {
    params.set("body", body);
  }

  const query = params.toString();
  return query ? `mailto:${to}?${query}` : `mailto:${to}`;
}

export function buildGmailWebComposeUrl({ to, subject, body }: EmailComposeParams) {
  const safeBody = body ?? "";
  return `https://mail.google.com/mail/?view=cm&fs=1&to=${encodeURIComponent(to)}&su=${encodeURIComponent(
    subject
  )}&body=${encodeURIComponent(safeBody)}`;
}

export function buildGmailIOSAppComposeUrl({ to, subject, body }: EmailComposeParams) {
  const params = new URLSearchParams();
  params.set("to", to);
  params.set("subject", subject);
  if (body) {
    params.set("body", body);
  }

  return `googlegmail://co?${params.toString()}`;
}

export function buildGmailAndroidIntentComposeUrl({ to, subject, body }: EmailComposeParams) {
  const params = new URLSearchParams();
  params.set("to", to);
  params.set("subject", subject);
  if (body) {
    params.set("body", body);
  }

  return `intent://co?${params.toString()}#Intent;scheme=googlegmail;package=com.google.android.gm;end`;
}
