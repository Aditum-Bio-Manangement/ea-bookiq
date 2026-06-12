import type { Metadata } from 'next'
import { Geist, Geist_Mono } from 'next/font/google'
import { Analytics } from '@vercel/analytics/next'
import './globals.css'

const _geist = Geist({ subsets: ["latin"] });
const _geistMono = Geist_Mono({ subsets: ["latin"] });

export const metadata: Metadata = {
  title: 'AB Room IQ - Aditum Bio',
  description: 'Intelligent conference room booking for Aditum Bio employees. An Outlook add-in by Aditum Bio.',
  generator: 'Developed by Caleb Klobe - Aditum Bio',
  icons: {
    icon: '/images/favicon.ico',
    apple: '/icons/icon-64.png',
  },
}

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode
}>) {
  return (
    <html lang="en">
      <body className="font-sans antialiased">
        {children}
        <Analytics />
      </body>
    </html>
  )
}
