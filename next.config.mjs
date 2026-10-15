/** @type {import('next').NextConfig} */

if (!process.env.NEXT_PUBLIC_REDIRECT_URI) {
  throw new Error(
    "Missing NEXT_PUBLIC_REDIRECT_URI. Set it to the hosted task pane URL before build or start. For Azure Static Web Apps use https://zealous-tree-04263bb0f.5.azurestaticapps.net/taskpane"
  )
}

const nextConfig = {
  output: "export",
  trailingSlash: false,
  typescript: {
    ignoreBuildErrors: true,
  },
  images: {
    unoptimized: true,
  },
}

export default nextConfig
