class GmailAutocomplete < Formula
  desc "Gmail to Outlook Autocomplete Builder"
  homepage "https://github.com/yourusername/gmail-autocomplete"
  version "1.0.0"
  
  # For actual deployment, you'd host the binary and update this URL
  url "file://#{Dir.pwd}/dist/gmail-autocomplete"
  sha256 "PLACEHOLDER_SHA256"
  
  def install
    bin.install "gmail-autocomplete"
  end
  
  test do
    system "#{bin}/gmail-autocomplete", "--help"
  end
end
