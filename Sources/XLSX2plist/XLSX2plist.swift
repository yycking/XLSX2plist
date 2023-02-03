import Foundation
import CoreXLSX

public struct XLSX2plist {
    let sheets: [String: [[String: String?]]]
    
    public init(file: URL) {
        let file = XLSXFile(filepath: file.path(percentEncoded: false))!
        let strings = try! file.parseSharedStrings()!
        let book = try! file.parseWorkbooks().first!
        sheets = try! file.parseWorksheetPathsAndNames(workbook: book)
            .reduce(into: [:]) { result, item in
                guard let name = item.name else {return}
                let sheet = try! file.parseWorksheet(at: item.path)
                var rows = sheet.data!.rows
                    .map { row -> [(key: String, value: String?)] in
                        row.cells.map { cell -> (key: String, value: String?) in
                            let key = cell.reference.column.value
                            let value = cell.stringValue(strings)
                            return (key: key, value: value)
                        }
                    }
                let keys = rows.removeFirst()
                    .reduce(into: [:]) { result, cell in
                        result[cell.key] = cell.value
                    }
                    
                result[name] = rows
                    .map { row -> [String : String?] in
                        row.reduce(into: [:]) { result, cell in
                            guard let key = keys[cell.key] else {return}
                            let value = cell.value
                            result[key] = value
                        }
                    }
            }
    }
    
    public func write(to path: URL) {
        sheets.forEach { (name: String, values: [[String: String?]]) in
            let fileURL = path.appendingPathComponent("\(name).plist")
            do {
                try PropertyListEncoder().encode(values)
                    .write(to: fileURL)
            } catch {
                print(error)
            }
        }
    }
}
