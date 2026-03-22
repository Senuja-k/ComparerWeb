import Link from "next/link";

export default function CosmeticsComingSoonPage() {
  return (
    <div className="min-h-screen bg-gradient-to-b from-[#f8fafc] to-[#f5f0ff] flex flex-col justify-center items-center text-center">
      <h1
        className="text-5xl font-bold m-0 mb-2"
        style={{
          background: "linear-gradient(90deg, #7c4dff, #e040fb)",
          WebkitBackgroundClip: "text",
          WebkitTextFillColor: "transparent",
        }}
      >
        💄 Cosmetics.lk
      </h1>
      <p className="text-lg text-[#787878] mb-8">Coming Soon</p>
      <Link href="/sales-report" className="text-[#7c4dff] no-underline text-[15px] hover:underline">
        ← Back to Sales Report
      </Link>
    </div>
  );
}
