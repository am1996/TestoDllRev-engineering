#define DAT_1000d8b0 0xA
void stack_probe_memory_management(void)
{
  uint in_EAX;
  undefined1 *puVar1;
  undefined4 unaff_retaddr;
  
  puVar1 = &stack0x00000004;
  for (; 0xfff < in_EAX; in_EAX = in_EAX - 0x1000) {
    puVar1 = puVar1 + -0x1000;
  }
  *(undefined4 *)(puVar1 + (-4 - in_EAX)) = unaff_retaddr;
  return;
}

void __thiscall create_sub_storage(void* pClassBase, IStorage* pRootStorage, wchar_t* pElementName)
{
    IStorage* pTempSourceStorage = nullptr;
    IStorage* pNewSubStorage = nullptr;
    HRESULT hr;
    int index = 0;
    
    // 1. LOOKUP: Find the object in the internal array that matches the name/ID
    int objectCount = *(int*)((int)pClassBase + 0x28);
    void** pObjectArray = *(void***)((int)pClassBase + 0x20);
    void** pInterfaceArray = *(void***)((int)pClassBase + 0x24);
    
    void* pFoundObject = nullptr;
    
    for (index = 0; index < objectCount; index++) {
        if ((wchar_t*)pObjectArray[index] == pElementName) {
            // Retrieve the functional interface for this specific data record
            pFoundObject = pInterfaceArray[index];
            break;
        }
    }

    if (pFoundObject == nullptr) return;

    // 2. PREPARE SOURCE: Ask the object to provide its data via a temporary storage
    // This is a virtual call to the object's own "GetStorage" method
    (*(void (__stdcall *)(void*, const char*, IStorage**))(*(int*)pFoundObject))(
        pFoundObject, "&DAT_1000d8d8", &pTempSourceStorage
    );

    if (pTempSourceStorage != nullptr) {
        wchar_t formattedName[32];
        // Create the name for the internal folder (e.g., based on the timestamp)
        swprintf(formattedName, L"%s", pElementName);

        // 3. CREATE SUB-FOLDER: Create a nested IStorage inside the root file
        // 0x14 is the Vtable offset for IStorage::CreateStorage
        hr = pRootStorage->lpVtbl->CreateStorage(
            pRootStorage, 
            formattedName, 
            STGM_READWRITE | STGM_SHARE_EXCLUSIVE | STGM_DIRECT, 
            0, 0, 
            &pNewSubStorage
        );

        if (hr == S_OK) {
            // 4. COMMIT DATA: Copy everything from the temp source to the new file location
            // 0x1c is the Vtable offset for IStorage::CopyTo
            pTempSourceStorage->lpVtbl->CopyTo(pTempSourceStorage, 0, NULL, NULL, pNewSubStorage);
        }

        // 5. CLEANUP: Release COM references to prevent memory leaks
        if (pNewSubStorage) pNewSubStorage->lpVtbl->Release(pNewSubStorage);
        if (pTempSourceStorage) pTempSourceStorage->lpVtbl->Release(pTempSourceStorage);
    }
}

undefined4 __thiscall Write_To_Vi2(void *this,LPCSTR filepath)
{
  int filePath_len;
  long ret_from_storage_open;
  HRESULT stream_opened;
  CString *this_00;
  undefined4 local_2c;
  undefined4 magic_numbers;
  undefined4 base_tick;
  ULONG local_20;
  int local_1c;
  IStorage *storage;
  IStream *stream;
  void *exception_list;
  undefined1 *puStack_c;
  int local_8;
  LPCSTR filepath_converted;
  
  filepath_converted = filepath;
  local_8 = 0xffffffff;
  puStack_c = &LAB_10008e18;
  exception_list = ExceptionList;
  if ((filepath != (LPCSTR)0x0) &&
     (ExceptionList = &exception_list, filePath_len = lstrlenA(filepath), filePath_len != 0)) {
    storage = (IStorage *)0x0;
    local_8 = 0;
    filePath_len = lstrlenA(filepath_converted);
    stack_probe_memory_management();
    MultiByteToWideChar(0,0,filepath,-1,(LPWSTR)&stack0xffffffc8,filePath_len + 1);
    ret_from_storage_open = TcOpenStorage((ushort *)&stack0xffffffc8,0x12,&storage,1);
    if (ret_from_storage_open == 0) {
      stream = (IStream *)0x0;
      local_8._0_1_ = 1;
      stream_opened =
           (*storage->lpVtbl->CreateStream)(storage,(OLECHAR *)&DAT_1000d8b0,0x1012,0,0,&stream);
      if (stream_opened == 0) {
        magic_numbers = 0x500;
        (*stream->lpVtbl->Write)(stream,&magic_numbers,2,&local_20);
        local_2c = 0;
        (*stream->lpVtbl->Write)(stream,&local_2c,2,&local_20);
        local_1c = *(int *)((int)this + 0x28);
        (*stream->lpVtbl->Write)(stream,&local_1c,4,&local_20);
        filePath_len = 0;
        if (0 < local_1c) {
          do {
            base_tick = *(undefined4 *)(*(int *)((int)this + 0x20) + filePath_len * 4);
            (*stream->lpVtbl->Write)(stream,&base_tick,4,&local_20);
            CString::CString((CString *)&filepath);
            local_8._0_1_ = 2;
            CString::Format(this_00,(char *)&filepath);
            Tcsputs(stream,filepath);
            local_8._0_1_ = 1;
            CString::~CString((CString *)&filepath);
            filePath_len = filePath_len + 1;
          } while (filePath_len < local_1c);
        }
        filePath_len = 0;
        if (0 < local_1c) {
          do {
            create_sub_storage(this,(int *)storage,
                         *(wchar_t **)(*(int *)((int)this + 0x20) + filePath_len * 4));
            filePath_len = filePath_len + 1;
          } while (filePath_len < local_1c);
        }
        local_8 = (uint)local_8._1_3_ << 8;
        if (stream != (IStream *)0x0) {
          (*stream->lpVtbl->Release)(stream);
        }
        local_8 = 0xffffffff;
        if (storage != (IStorage *)0x0) {
          (*storage->lpVtbl->Release)(storage);
        }
        ExceptionList = exception_list;
        return 1;
      }
      local_8 = (uint)local_8._1_3_ << 8;
      if (stream != (IStream *)0x0) {
        (*stream->lpVtbl->Release)(stream);
      }
    }
    local_8 = 0xffffffff;
    if (storage != (IStorage *)0x0) {
      (*storage->lpVtbl->Release)(storage);
    }
  }
  ExceptionList = exception_list;
  return 0;
}

