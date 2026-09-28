import "./globals.css";

export const metadata = {
  title: "AC-34 Notice & Hearing Dashboard",
  description: "SIR-2026 PS-wise AERO / Ad.AERO notice and hearing dashboard",
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return <html lang="en"><body>{children}</body></html>;
}
