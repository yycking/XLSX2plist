//
//  main.swift
//  
//
//  Created by Wayne Yeh on 2023/2/3.
//

import Foundation

print("run")
let downloadURL = FileManager.default
    .urls(for: .downloadsDirectory, in: .userDomainMask).first!
let xlsx = downloadURL.appendingPathComponent("file.xlsx")
XLSX2plist(file: xlsx).write(to: downloadURL)
