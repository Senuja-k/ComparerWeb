import Link from "next/link";

export default function SalesReportPage() {
  const cards = [
    {
      title: "SupplementVault.lk",
      icon: "💊",
      desc: "Generate monthly sales performance reports for SupplementVault merchants",
      href: "/supplement-vault",
      color: "#e040fb",
    },
    {
      title: "Cosmetics.lk",
      icon: "💄",
      desc: "Sales reporting for Cosmetics.lk merchants",
      href: "/cosmetics-coming-soon",
      color: "#7c4dff",
    },
  ];

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
          📈 Sales Report
        </h1>
        <p className="text-[15px] text-[#787878] mt-2">
          Select the platform to generate sales reports
        </p>
      </div>

      <div className="flex-1 flex justify-center items-center px-5 pb-5">
        <div className="grid grid-cols-2 gap-5 max-w-[800px] w-full">
          {cards.map((c) => (
            <Link
              key={c.href}
              href={c.href}
              className="bg-white rounded-2xl shadow-md p-8 flex items-center justify-between h-[140px] transition-transform hover:-translate-y-1 hover:shadow-lg no-underline text-inherit"
              style={{ border: `3px solid ${c.color}` }}
            >
              <div className="max-w-[75%]">
                <h2 className="text-xl font-bold m-0 mb-2 flex items-center gap-2">
                  <span>{c.icon}</span> {c.title}
                </h2>
                <p className="text-[#787878] m-0 text-sm leading-relaxed">{c.desc}</p>
              </div>
              <div className="text-[22px] font-bold text-[#7c4dff]">→</div>
            </Link>
          ))}
        </div>
      </div>

      <Link href="/" className="text-[#7c4dff] no-underline text-[15px] mb-8 hover:underline">
        ← Back to Dashboard
      </Link>
    </div>
  );
}
