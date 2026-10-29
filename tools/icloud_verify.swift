// icloud_verify.swift —— iCloud 云盘上传复查（总纲第十节第 11 步；2026-10-28 谷歌云盘停用后取代 drive_verify.py）。
//
// 用法：swift ~/flight-repos/pub/tools/icloud_verify.swift <文件或文件夹> [...]
//   对每个本地文件读系统的 iCloud 状态（ubiquitousItemIsUploaded）；文件夹逐个查其中的文件（不递归，跳过 . 开头的）。
//   全部已上传退出码 0；有未上传的列出来、退出码 1——过几分钟再跑。
//   一直不动时看 `brctl status com.apple.CloudDocs` 的上传队列；「无法访问 iCloud 账号」在大批排队时会暂时出现。
import Foundation

let keys: Set<URLResourceKey> = [.ubiquitousItemIsUploadedKey, .ubiquitousItemIsUploadingKey, .ubiquitousItemUploadingErrorKey]
let fm = FileManager.default
var files: [String] = []
for a in CommandLine.arguments.dropFirst() {
    var isDir: ObjCBool = false
    if !fm.fileExists(atPath: a, isDirectory: &isDir) { print("-- 本地没有：", a); continue }
    if isDir.boolValue {
        for n in ((try? fm.contentsOfDirectory(atPath: a)) ?? []).sorted() where !n.hasPrefix(".") {
            var d: ObjCBool = false
            let p = (a as NSString).appendingPathComponent(n)
            if fm.fileExists(atPath: p, isDirectory: &d), !d.boolValue { files.append(p) }
        }
    } else { files.append(a) }
}
var pending = 0
for p in files {
    let v = try? URL(fileURLWithPath: p).resourceValues(forKeys: keys)
    if v?.ubiquitousItemIsUploaded == true { continue }
    pending += 1
    let st = v?.ubiquitousItemIsUploading == true ? "上传中" : "未上传"
    print("\(st)\t\(v?.ubiquitousItemUploadingError?.localizedDescription ?? "")\t\(p)")
}
print("共 \(files.count) 个，已上传 \(files.count - pending)，未确认 \(pending)")
exit(pending == 0 ? 0 : 1)
