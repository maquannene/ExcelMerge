//
//  ViewController.m
//  ExcelMergeTool
//
//  Created by 马权 on 05/03/2017.
//  Copyright © 2017 马权. All rights reserved.
//

#import "ViewController.h"
#import <LibXL/LibXL.h>
#import "Masonry/Masonry.h"

static NSString * const kChooseInputFile = @"Choose Input File";
static NSString * const kChooseOutputFile = @"Choose Output File";


static NSString * const kSpecialKeySN = @"Sheet Name";
static NSString * const kSpecialKeyBRC = @"BATT_REWORK_CNT";

static NSString * const kSpecialKeyFTC = @"Family Type Code";
static NSString * const kSpecialKeyFTC_col = @"Reported Failures";

static NSString * const kSpecialKeyHSG = @"HSG";
static NSString * const kSpecialKeyHSG_col = @"Timestamp";


@interface ViewController () <NSOpenSavePanelDelegate, NSTextFieldDelegate>

{
    __weak IBOutlet NSButton *_chooseInputBtn;
    __weak IBOutlet NSButton *_chooseOutputBtn;
    __strong NSTextField *_inputTextField;
    __strong NSTextField *_outputTextField;
    
    __strong NSOpenPanel *_openPanel;
    NSArray<NSURL *> *_inputFileURLs;
    NSURL *_outFileURL;
}

@end

@implementation ViewController

- (void)viewDidLoad {
    [super viewDidLoad];
    
    _inputTextField = [[NSTextField alloc] initWithFrame:CGRectZero];
    _inputTextField.editable = NO;
    _inputTextField.selectable = YES;
    
    [self.view addSubview:_inputTextField];
    [_inputTextField mas_makeConstraints:^(MASConstraintMaker *make) {
        make.top.equalTo(_chooseInputBtn);
        make.bottom.lessThanOrEqualTo(_chooseOutputBtn.mas_top).offset(-20);
        make.left.equalTo(_chooseInputBtn.mas_right).offset(20);
        make.right.equalTo(self.view).offset(-20);
    }];
    
    _outputTextField = [[NSTextField alloc] initWithFrame:CGRectZero];
    _outputTextField.editable = NO;
    _outputTextField.selectable = YES;
    [self.view addSubview:_outputTextField];
    [_outputTextField mas_makeConstraints:^(MASConstraintMaker *make) {
        make.top.equalTo(_chooseOutputBtn);
        make.left.equalTo(_chooseOutputBtn.mas_right).offset(20);
        make.right.equalTo(self.view).offset(-20);
    }];
}

- (IBAction)chooseInputAction:(id)sender {
    _openPanel = [NSOpenPanel openPanel];
    _openPanel.title = kChooseInputFile;
    _openPanel.allowsMultipleSelection = YES;
    _openPanel.canChooseFiles = YES;
    _openPanel.canChooseDirectories = NO;
    _openPanel.delegate = self;
    [_openPanel beginWithCompletionHandler:^(NSInteger result) {
        
    }];
}

- (IBAction)chooseOutputAction:(id)sender {
    _openPanel = [NSOpenPanel openPanel];
    _openPanel.title = kChooseOutputFile;
    _openPanel.allowsMultipleSelection = NO;
    _openPanel.canChooseFiles = YES;
    _openPanel.canChooseDirectories = YES;
    _openPanel.delegate = self;
    [_openPanel beginWithCompletionHandler:^(NSInteger result) {
        
    }];
}

- (BOOL)panel:(id)sender validateURL:(NSURL *)url error:(NSError **)outError
{
    NSOpenPanel *openPanel = nil;
    if ([sender isKindOfClass:[NSOpenPanel class]]) {
        openPanel = (NSOpenPanel *)sender;
    }
    if ([openPanel.title isEqualToString:kChooseInputFile]) {
        _inputFileURLs = openPanel.URLs;
        NSMutableString *paths = [NSMutableString new];
        [_inputFileURLs enumerateObjectsUsingBlock:^(NSURL * _Nonnull obj, NSUInteger idx, BOOL * _Nonnull stop) {
            [paths appendString:[NSString stringWithFormat:@"%@\n", obj.path]];
        }];
        _inputTextField.stringValue = paths;
    }
    if ([openPanel.title isEqualToString:kChooseOutputFile]) {
        _outFileURL = openPanel.URLs.firstObject;
        _outputTextField.stringValue = _outFileURL.path;
    }
    return YES;
}

- (IBAction)readExcel:(id)sender {
    
    if (_inputFileURLs.count == 0 ||
        _outFileURL.path.length == 0) {
        return;
    }
    
    //  输出 BookHandle
    BookHandle outputBookHandle = [[self class] bookHandleFilePath:_outFileURL.path];
    SheetHandle outputSheetHandle = xlBookGetSheet(outputBookHandle, 0);
    
    //  读取输出 keyArray
    NSMutableArray<NSString *> *keysArray = [[self class] titleArrayWithSheetHandle:outputSheetHandle];

    //  遍历多个 url
    [_inputFileURLs enumerateObjectsUsingBlock:^(NSURL * _Nonnull obj, NSUInteger idx, BOOL * _Nonnull stop) {
        //  读取 book 的基础数据
        BookHandle tInputBookHandle = [[self class] bookHandleFilePath:obj.path];
        NSMutableArray<NSDictionary *> *sheetsDicArray = @[].mutableCopy;
        int sheetCount = xlBookSheetCount(tInputBookHandle);
        xlBookRelease(tInputBookHandle);    //  这里用完就关掉
        
        //  遍历 book 中每一个 sheet
        for (int sheetIndex = 0; sheetIndex < sheetCount; sheetIndex++) {
            
            //  由于试用版 sdk 有问题，每个 book 只能读出来 3 个 sheet，所以每次读取 sheet 都重新打开 book
            BookHandle bookHandle = [[self class] bookHandleFilePath:obj.path];
            SheetHandle sheetHandle = xlBookGetSheet(bookHandle, sheetIndex);
            NSString *sheetName = [[self class] sheetNameWithSheetHandle:sheetHandle];
            
            int sheetRowCount = xlSheetLastRow(sheetHandle);
            BOOL beginSearchHSG = NO;
            
            //  数据
            NSMutableDictionary *sheetDic = @{}.mutableCopy;
            sheetDic[kSpecialKeySN] = @{@"String" : sheetName};
            
            for (int row = 0; row < sheetRowCount; row++) {
                //  读取第 0 行的所有 key
                const char *cKey = xlSheetReadStr(sheetHandle, row, 0, NULL);
                if (cKey != NULL) {
                    //  去掉 ：
                    NSString *key = [NSString stringWithUTF8String:cKey];
                    if ([[key substringFromIndex:key.length - 1] isEqualToString:@":"]) {
                        key = [key substringToIndex:key.length - 1];
                    }
                    //  存在于 keysArray 的 key，就读取 value，加入到 dic 中
                    if ([keysArray containsObject:key]) {
                        const char *cValue = NULL;
                        //  特殊处理
                        if ([key isEqualToString:kSpecialKeyBRC]) {
                            cValue = xlSheetReadStr(sheetHandle, row, 2, NULL);
                        }
                        else {
                            cValue = xlSheetReadStr(sheetHandle, row, 1, NULL);
                        }
                        
                        if (cValue != NULL) {
                            NSString *value = [NSString stringWithUTF8String:cValue];
                            sheetDic[key] = @{@"String" : value}.mutableCopy;
                        }

                    }

                    //  special FTC
                    if ([key isEqualToString:kSpecialKeyFTC_col]) {
                        NSString *value = [[self class] stringWithSheetHandle:sheetHandle row:row col:3];
                        if (value) {
                            sheetDic[kSpecialKeyFTC] = @{@"String" : value};
                        }
                    }
                    
                    
                    //  special HSG
                    if (beginSearchHSG) {
                        NSString *key = [[self class] stringWithSheetHandle:sheetHandle row:row col:8];
                        if ([key containsString:kSpecialKeyHSG]) {
                            NSString *value = [[self class] stringWithSheetHandle:sheetHandle row:row col:3];
                            if (value) {
                                sheetDic[kSpecialKeyHSG] = @{@"String" : value};
                            }
                        }
                    }
                    
                    if ([key containsString:kSpecialKeyHSG_col]) {
                        beginSearchHSG = YES;
                    }
                }
            }
            
            //  dicInfo 塞入 sheetsDicArray
            if (sheetDic.allKeys.count > 0) {
                [sheetsDicArray addObject:sheetDic];
            }
            
            xlBookRelease(bookHandle);
        }
        
        //  将一个 book 中提取的数据，写入 outputSheetHandle 中
        __block int rowCount = xlSheetLastRow(outputSheetHandle);
        [sheetsDicArray enumerateObjectsUsingBlock:^(NSDictionary * _Nonnull keyValueDic, NSUInteger idx, BOOL * _Nonnull stop) {
            [keysArray enumerateObjectsUsingBlock:^(NSString * _Nonnull key, NSUInteger col, BOOL * _Nonnull stop) {
                if ([[key substringFromIndex:key.length - 1] isEqualToString:@":"]) {
                    key = [key substringToIndex:key.length - 1];
                }
                NSDictionary *value = keyValueDic[key];
                if (value[@"String"]) {
                    xlSheetWriteStr(outputSheetHandle, rowCount, (int)col, [value[@"String"] UTF8String], NULL);
                }
//                if (value[@"Format"]) {
//                    FormatHandle format;
//                    [value[@"Format"] getValue:&format];
//                    xlSheetSetCellFormat(outputSheetHandle, rowCount, (int)col + 1, format);
//                }
            }];
            rowCount++;
        }];
    }];

    [[self class] saveBookHandle:outputBookHandle withFlePath:_outFileURL.path];
    
    xlBookRelease(outputBookHandle);
}

+ (NSString *)sheetNameWithSheetHandle:(SheetHandle)sheetHandle
{
    const char* cSheetName = xlSheetName(sheetHandle);
    return [NSString stringWithUTF8String:cSheetName];
}
    
+ (NSString *)stringWithSheetHandle:(SheetHandle)sheetHandle
                                row:(int)row
                                col:(int)col
{
    const char *cValue = xlSheetReadStr(sheetHandle, row, col, NULL);
    if (cValue != NULL) {
        return [NSString stringWithUTF8String:cValue];
    }
    return nil;
}

+ (NSMutableArray *)titleArrayWithSheetHandle:(SheetHandle)sheetHandle
{
    NSMutableArray *titleArray = @[].mutableCopy;
    int colCount = xlSheetLastColA(sheetHandle);
    for (int col = 0; col < colCount; col++) {
        const char * cTitle = xlSheetReadStr(sheetHandle, 1, col, NULL);
        if (cTitle != NULL) {
            NSString *title = [NSString stringWithUTF8String:cTitle];
            [titleArray addObject:title];
        }
    }
    return titleArray;
}

+ (void)saveBookHandle:(BookHandle)bookhandle
           withFlePath:(NSString *)filePath
{
    NSDateFormatter *dateFormat = [[NSDateFormatter alloc] init];
    [dateFormat setDateFormat:@"yyyyMMddhhmmss"];
    NSString *dateString = [dateFormat stringFromDate:[NSDate date]];
    NSString *path = [NSString stringWithFormat:@"%@_%@.%@", filePath.stringByDeletingPathExtension, dateString, filePath.pathExtension];
    xlBookSaveA(bookhandle, [path UTF8String]);
}

+ (BookHandle)bookHandleFilePath:(NSString *)filePath
{
    BookHandle handle = NULL;
    NSString *extension = [filePath pathExtension];
    if ([extension isEqualToString:@"xlsx"]) {
        handle = xlCreateXMLBook();
    }
    if ([extension isEqualToString:@"xls"]) {
        handle = xlCreateBook();
    }
    
    xlBookLoad(handle, [filePath UTF8String]);
    return handle;
}

@end
