#pragma once
// Every first-party ASan Debug translation unit must carry real compiler instrumentation.
#ifndef __SANITIZE_ADDRESS__
#error ASan Debug requires compiler AddressSanitizer instrumentation.
#endif
