using System; using System.Text; using Ozbekca;
class T {
  // Ekran (matn maydoni) simulyatsiyasi
  static string Type(Engine e, string keys) {
    var sb = new StringBuilder();
    foreach (char c in keys) {
      if (c == '\b') { e.Backspace(); if (sb.Length>0) sb.Length--; continue; }
      if (c == '|') { e.Reset(); continue; } // sichqoncha bosildi
      var n = e.OnChar(c);
      if (!n.Yutish) { sb.Append(c); continue; }
      sb.Length = Math.Max(0, sb.Length - n.Ochirish);
      sb.Append(n.Yozish);
    }
    return sb.ToString();
  }
  static int fail = 0;
  static void Ok(string keys, string want) {
    var got = Type(new Engine(), keys);
    if (got != want) { fail++; Console.WriteLine("XATO: [" + keys + "] -> [" + got + "] kutilgan [" + want + "]"); }
    else Console.WriteLine("ok  : " + keys + " -> " + got);
  }
  static void Main() {
    Ok("O'qituvchi ", "Öqituvçi ");
    Ok("Shahar chiroyli. ", "Şahar çiroyli. ");
    Ok("SHAHAR ", "ŞAHAR ");
    Ok("tog'ga bordik", "toğga bordik");
    Ok("o`zbek o‘zbek oʻzbek", "özbek özbek özbek");
    Ok("Photoshop ", "Photoshop ");
    Ok("Chevrolet, ", "Chevrolet, ");
    Ok("English ", "English ");
    Ok("show ", "show ");
    Ok("o'clock ", "o'clock ");
    Ok("O'Brien.", "O'Brien.");
    Ok("Ishoqjon ", "Ishoqjon ");
    Ok("SHA-256 ", "SHA-256 ");
    Ok("shh", "sh");
    Ok("o''", "o'");
    Ok("s|h", "sh");           // sichqoncha bosilgach — o'girilmaydi
    Ok("maktab\b\b\bab", "makab");
    Ok("chashma\n", "çaşma\n");
    Ok("ta'lim ", "ta'lim ");
    Ok("Toshkent\tqo'shiq", "Toşkent\tqöşiq");
    Console.WriteLine(fail == 0 ? "HAMMASI OK" : fail + " ta xato");
    Environment.Exit(fail);
  }
}
