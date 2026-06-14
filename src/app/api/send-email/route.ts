import { NextResponse } from 'next/server';
import nodemailer from 'nodemailer';

export async function POST(request: Request) {
  try {
    const { recipients } = await request.json();

    if (!recipients || !Array.isArray(recipients) || recipients.length === 0) {
      return NextResponse.json({ error: 'No recipients provided' }, { status: 400 });
    }

    const { emailTemplateHtml } = await import('@/lib/emailTemplate');

    const encoder = new TextEncoder();

    const stream = new ReadableStream({
      async start(controller) {
        try {
          // Create a transporter using Gmail settings
          const transporter = nodemailer.createTransport({
            service: 'gmail',
            auth: {
              user: process.env.EMAIL_USER,
              pass: process.env.EMAIL_PASS,
            },
          });

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
              
              const progressData = JSON.stringify({ 
                progress: sentCount, 
                total: recipients.length,
                email: recipient
              });
              controller.enqueue(encoder.encode(`data: ${progressData}\n\n`));
            } catch (err) {
              console.error(`Failed to send email to ${recipient}:`, err);
              // Report progress even if sending failed for this recipient, so the total matches
              const errorData = JSON.stringify({ 
                progress: sentCount, // kept at successful sends
                total: recipients.length,
                email: recipient,
                error: true
              });
              controller.enqueue(encoder.encode(`data: ${errorData}\n\n`));
            }
          }

          controller.enqueue(encoder.encode(`data: ${JSON.stringify({ done: true, sentCount, total: recipients.length })}\n\n`));
          controller.close();
        } catch (error) {
          console.error('Error in stream:', error);
          controller.enqueue(encoder.encode(`data: ${JSON.stringify({ error: 'Failed to setup email transport' })}\n\n`));
          controller.close();
        }
      }
    });

    return new Response(stream, {
      headers: {
        'Content-Type': 'text/event-stream',
        'Cache-Control': 'no-cache',
        'Connection': 'keep-alive',
      },
    });

  } catch (error) {
    console.error('Error sending emails:', error);
    return NextResponse.json({ error: 'Failed to process request' }, { status: 500 });
  }
}
