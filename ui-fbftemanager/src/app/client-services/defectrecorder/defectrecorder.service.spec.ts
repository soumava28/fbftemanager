import { TestBed } from '@angular/core/testing';

import { DefectrecorderService } from './defectrecorder.service';

describe('DefectrecorderService', () => {
  beforeEach(() => TestBed.configureTestingModule({}));

  it('should be created', () => {
    const service: DefectrecorderService = TestBed.get(DefectrecorderService);
    expect(service).toBeTruthy();
  });
});
