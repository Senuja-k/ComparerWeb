import Link from "next/link";

const modes = [
  {
    key: "created-date",
    href: "/supplement-vault/created-date",
    icon: "📅",
    title: "Created Date",
    description: "Generate sales report filtered by order created date",
  },
  {
    key: "fulfilled-date",
    href: "/supplement-vault/fulfilled-date",
    icon: "✅",
    title: "Fulfilled Date",
    description: "Generate sales report filtered by order fulfilled date",
  },
];

export default function SupplementVaultPage() {
  return (
    <div className="min-h-screen bg-gradient-to-b from-[#f8fafc] to-[#f5f0ff] flex flex-col items-center">
      <div className="pt-10 pb-5 text-center">
        <h1
          className="text-4xl font-bold m-0"
          style={{
            background: "linear-gradient(90deg, #e040fb, #7c4dff)",
            WebkitBackgroundClip: "text",
            WebkitTextFillColor: "transparent",
          }}
        >
          💊 SupplementVault.lk
        </h1>
        <p className="text-[15px] text-[#787878] mt-2">
          Select the date mode for your sales report
        </p>
      </div>

      <div className="flex-1 flex justify-center items-center px-5 pb-5">
        <div className="grid grid-cols-2 gap-5 max-w-[700px] w-full">
          {modes.map((m) => (
            <Link
              key={m.key}
              href={m.href}
              className="bg-white rounded-2xl shadow-md p-8 flex items-center justify-between h-[140px] transition-transform hover:-translate-y-1 hover:shadow-lg no-underline text-inherit"
              style={{ border: "3px solid #e040fb" }}
            >
              <div className="max-w-[75%]">
                <h2 className="text-xl font-bold m-0 mb-2 flex items-center gap-2 text-[#1e1e1e]">
                  <span>{m.icon}</span> {m.title}
                </h2>
                <p className="text-[#787878] m-0 text-sm leading-relaxed">{m.description}</p>
              </div>
              <div className="text-[22px] font-bold text-[#e040fb]">→</div>
            </Link>
          ))}
        </div>
      </div>

      <Link href="/sales-report" className="text-[#7c4dff] no-underline text-[15px] mb-8 hover:underline">
        ← Back to Sales Report
      </Link>
    </div>
  );
}
