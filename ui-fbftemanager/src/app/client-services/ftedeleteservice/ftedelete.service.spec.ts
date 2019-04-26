import { TestBed } from '@angular/core/testing';

import { FtedeleteService } from './ftedelete.service';

describe('FtedeleteService', () => {
  beforeEach(() => TestBed.configureTestingModule({}));

  it('should be created', () => {
    const service: FtedeleteService = TestBed.get(FtedeleteService);
    expect(service).toBeTruthy();
  });
});
