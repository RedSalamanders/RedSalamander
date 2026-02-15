#pragma once

#include "FolderWindow.h"

class FileOperationsIssuesPane final
{
public:
    static HWND Create(FolderWindow::FileOperationState* fileOps, FolderWindow* folderWindow, HWND ownerWindow) noexcept;

private:
    FileOperationsIssuesPane() = delete;
};
