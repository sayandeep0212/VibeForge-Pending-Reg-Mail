import { NextResponse } from 'next/server';
import nodemailer from 'nodemailer';

export async function POST(request: Request) {
  try {
    const { recipients } = await request.json();

    if (!recipients || !Array.isArray(recipients) || recipients.length === 0) {
      return NextResponse.json({ error: 'No recipients provided' }, { status: 400 });
    }

    // Create a transporter using Gmail settings
    const transporter = nodemailer.createTransport({
      service: 'gmail',
      auth: {
        user: process.env.EMAIL_USER,
        pass: process.env.EMAIL_PASS,
      },
    });

    const { emailTemplateHtml } = await import('@/lib/emailTemplate');

    // Send email to each recipient individually to ensure they are in the "To" field
    let sentCount = 0;
    for (const recipient of recipients) {
      try {
        const mailOptions = {
          from: process.env.EMAIL_USER,
          to: recipient,
          subject: 'Action Required: Complete Your VibeForge 1.0 Registration',
          text: 'Hello,\n\nYou have successfully created your profile for VibeForge 1.0, but your registration is still incomplete. Please complete your registration by registering your team.\n\nThank you!',
          html: emailTemplateHtml,
        };

        await transporter.sendMail(mailOptions);
        sentCount++;
      } catch (err) {
        console.error(`Failed to send email to ${recipient}:`, err);
      }
    }

    console.log(`Successfully sent ${sentCount} out of ${recipients.length} emails.`);

    return NextResponse.json({ success: true, message: 'Emails sent successfully' });
  } catch (error) {
    console.error('Error sending emails:', error);
    return NextResponse.json({ error: 'Failed to send emails' }, { status: 500 });
  }
}
