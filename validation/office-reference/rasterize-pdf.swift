import Foundation
import PDFKit
import AppKit
for file in CommandLine.arguments.dropFirst() {
 guard let doc = PDFDocument(url: URL(fileURLWithPath:file)) else {continue}
 print("\(file): \(doc.pageCount) pages")
 for i in 0..<doc.pageCount {
  let p = doc.page(at:i)!
  let r = p.bounds(for:.mediaBox)
  let w = Int(ceil(r.width*4/3)), h = Int(ceil(r.height*4/3))
  let color = CGColorSpaceCreateDeviceRGB()
  let c = CGContext(data:nil,width:w,height:h,bitsPerComponent:8,bytesPerRow:w*4,space:color,bitmapInfo:CGImageAlphaInfo.premultipliedLast.rawValue)!
  c.setFillColor(CGColor(gray:1,alpha:1));c.fill(CGRect(x:0,y:0,width:w,height:h))
  c.scaleBy(x:4/3,y:4/3);p.draw(with:.mediaBox,to:c)
  let image = NSBitmapImageRep(cgImage:c.makeImage()!)
  try! image.representation(using:.png,properties:[:])!.write(to:URL(fileURLWithPath:"\(file).\(i).png"))
  print("page \(i): \(w)x\(h) text: \(p.string ?? "")")
 }
}
