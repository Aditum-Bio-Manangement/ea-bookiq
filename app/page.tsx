import Link from "next/link";
import Image from "next/image";
import { Building2, Calendar, Users, Zap, CheckCircle, ArrowRight } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";

export default function HomePage() {
  return (
    <main className="min-h-screen bg-background">
      {/* Hero Section */}
      <section className="relative overflow-hidden border-b">
        <div className="absolute inset-0 bg-gradient-to-br from-[#2563eb]/5 via-transparent to-transparent" />
        <div className="relative max-w-5xl mx-auto px-6 py-24">
          <div className="flex flex-col items-center text-center">
            <Image
              src="/images/aditum-logo.png"
              alt="Aditum Bio"
              width={200}
              height={133}
              className="mb-8"
              priority
            />
            <div className="flex items-center gap-2 px-3 py-1 rounded-full bg-[#2563eb]/10 text-[#2563eb] text-sm font-medium mb-6">
              <Building2 className="size-4" />
              Outlook Add-in for Conference Room Booking
            </div>
            <h1 className="text-4xl md:text-5xl font-bold tracking-tight text-[#1e3a5f] max-w-3xl text-balance">
              AB Book IQ
            </h1>
            <p className="mt-4 text-lg text-muted-foreground max-w-2xl text-balance">
              Intelligent conference room booking for Aditum Bio employees. Automatically
              identifies your office location and shows available rooms for your meeting time.
            </p>
            <div className="mt-8 flex flex-wrap gap-4 justify-center">
              <Button asChild size="lg">
                <Link href="/taskpane">
                  Open Task Pane Preview
                  <ArrowRight className="ml-2 size-4" />
                </Link>
              </Button>
              <Button variant="outline" size="lg" asChild>
                <a href="/manifest.xml" download>
                  Download Manifest
                </a>
              </Button>
            </div>
          </div>
        </div>
      </section>

      {/* Features Section */}
      <section className="max-w-5xl mx-auto px-6 py-20">
        <div className="text-center mb-12">
          <h2 className="text-2xl font-bold text-foreground">How It Works</h2>
          <p className="mt-2 text-muted-foreground">
            Streamline conference room booking for Aditum Bio employees
          </p>
        </div>

        <div className="grid md:grid-cols-3 gap-6">
          <Card>
            <CardHeader>
              <div className="size-10 rounded-lg bg-primary/10 flex items-center justify-center mb-2">
                <Building2 className="size-5 text-primary" />
              </div>
              <CardTitle className="text-lg">Auto Office Detection</CardTitle>
              <CardDescription>
                Automatically identifies your office location based on
                security group membership (Cambridge or Oakland).
              </CardDescription>
            </CardHeader>
          </Card>

          <Card>
            <CardHeader>
              <div className="size-10 rounded-lg bg-primary/10 flex items-center justify-center mb-2">
                <Calendar className="size-5 text-primary" />
              </div>
              <CardTitle className="text-lg">Real-Time Availability</CardTitle>
              <CardDescription>
                Queries Microsoft Graph for live free/busy status during
                your exact meeting time window.
              </CardDescription>
            </CardHeader>
          </Card>

          <Card>
            <CardHeader>
              <div className="size-10 rounded-lg bg-primary/10 flex items-center justify-center mb-2">
                <Zap className="size-5 text-primary" />
              </div>
              <CardTitle className="text-lg">One-Click Booking</CardTitle>
              <CardDescription>
                Insert the best-fit room into your meeting with one click.
                Appears immediately in Scheduling Assistant.
              </CardDescription>
            </CardHeader>
          </Card>
        </div>
      </section>

      {/* Setup Section */}
      <section className="bg-muted/50 border-y">
        <div className="max-w-5xl mx-auto px-6 py-20">
          <div className="text-center mb-12">
            <h2 className="text-2xl font-bold text-foreground">Setup Requirements</h2>
            <p className="mt-2 text-muted-foreground">
              What you need to deploy AB Book IQ
            </p>
          </div>

          <div className="grid md:grid-cols-2 gap-8">
            <Card>
              <CardHeader>
                <CardTitle className="text-lg">Microsoft Entra App Registration</CardTitle>
              </CardHeader>
              <CardContent className="space-y-3">
                <div className="flex items-start gap-3">
                  <CheckCircle className="size-5 text-primary shrink-0 mt-0.5" />
                  <span className="text-sm text-muted-foreground">
                    Register a Single-Page Application (SPA)
                  </span>
                </div>
                <div className="flex items-start gap-3">
                  <CheckCircle className="size-5 text-primary shrink-0 mt-0.5" />
                  <span className="text-sm text-muted-foreground">
                    Configure redirect URIs for your hosted domain
                  </span>
                </div>
                <div className="flex items-start gap-3">
                  <CheckCircle className="size-5 text-primary shrink-0 mt-0.5" />
                  <span className="text-sm text-muted-foreground">
                    Grant admin consent for Graph permissions
                  </span>
                </div>
              </CardContent>
            </Card>

            <Card>
              <CardHeader>
                <CardTitle className="text-lg">Required Graph Permissions</CardTitle>
              </CardHeader>
              <CardContent className="space-y-3">
                <div className="flex items-start gap-3">
                  <Users className="size-5 text-muted-foreground shrink-0 mt-0.5" />
                  <span className="text-sm text-muted-foreground">
                    <strong>User.Read</strong> - Read user profile
                  </span>
                </div>
                <div className="flex items-start gap-3">
                  <Calendar className="size-5 text-muted-foreground shrink-0 mt-0.5" />
                  <span className="text-sm text-muted-foreground">
                    <strong>Calendars.Read.Shared</strong> - Check room availability
                  </span>
                </div>
                <div className="flex items-start gap-3">
                  <Building2 className="size-5 text-muted-foreground shrink-0 mt-0.5" />
                  <span className="text-sm text-muted-foreground">
                    <strong>Place.Read.All</strong> - Read room inventory
                  </span>
                </div>
                <div className="flex items-start gap-3">
                  <Users className="size-5 text-muted-foreground shrink-0 mt-0.5" />
                  <span className="text-sm text-muted-foreground">
                    <strong>GroupMember.Read.All</strong> - Read group membership
                  </span>
                </div>
              </CardContent>
            </Card>
          </div>
        </div>
      </section>

      {/* Environment Variables Section */}
      <section className="max-w-5xl mx-auto px-6 py-20">
        <div className="text-center mb-12">
          <h2 className="text-2xl font-bold text-foreground">Environment Variables</h2>
          <p className="mt-2 text-muted-foreground">
            Configure these variables in your Vercel project settings
          </p>
        </div>

        <Card>
          <CardContent className="p-6">
            <div className="font-mono text-sm space-y-2 bg-muted p-4 rounded-lg">
              <p><span className="text-muted-foreground"># Azure AD App Registration</span></p>
              <p>NEXT_PUBLIC_AZURE_CLIENT_ID=your-app-client-id</p>
              <p>NEXT_PUBLIC_AZURE_TENANT_ID=your-tenant-id</p>
              <p>NEXT_PUBLIC_REDIRECT_URI=https://your-domain.com/taskpane</p>
            </div>
          </CardContent>
        </Card>
      </section>

      {/* Footer */}
      <footer className="border-t">
        <div className="max-w-5xl mx-auto px-6 py-8">
          <div className="flex flex-col md:flex-row justify-between items-center gap-4">
            <div className="flex items-center gap-3">
              <Image
                src="/images/aditum-logo.png"
                alt="Aditum Bio"
                width={100}
                height={67}
                className="opacity-80"
              />
              <p className="text-sm text-muted-foreground">
                AB Book IQ - Conference Room Booking
              </p>
            </div>
            <div className="flex gap-4">
              <Link href="/taskpane" className="text-sm text-muted-foreground hover:text-foreground">
                Task Pane Preview
              </Link>
              <a href="/manifest.xml" className="text-sm text-muted-foreground hover:text-foreground">
                Manifest
              </a>
            </div>
          </div>
        </div>
      </footer>
    </main>
  );
}
